// mailer.go —— 本地文件内容汇报邮件发送器
//
// 功能概览：
//   1) 命令行参数：-c 配置文件、-f 本地文件、-b 正文、-s 主题、-x 文本导出、
//      --help(中文)、--dry-run、--verbose 等
//   2) 支持的本地文件：
//        a) 表格 .xlsx/.csv/.tsv（含内置 xlsx 解析、日期序列号转真实日期、比率列按百分比展示）；
//        b) Office 文档 .docx/.pptx（zip+XML 自动抽取文字）；
//        c) 文本 .txt/.md/.log/.text。
//   3) 两种用法：
//        a) AI + Skill 协作：AI 用 “mailer -x -f <文件>” 取文本(或自行读文件)→ 总结成一段话与一句话主题
//           → 调用 “mailer -c config.json -s <主题> -b <正文>” 发送；
//        b) 手工运行：无 -b 时由脚本对文件做确定性自动摘要兜底并给出提示。
//   4) 通过企业微信邮箱（腾讯企业邮箱 smtp.exmail.qq.com:465，SSL）发给
//      默认收件人（由 config defaults.to 或 -t 指定），不抄送；主题默认概括文件内容(可用 -s 或配置覆盖)。
//
// 设计约束：
//   * 零第三方依赖，仅用 Go 标准库（xlsx/docx/pptx 均为内置 zip/XML 解析，覆盖常见导出），离线可构建；
//   * 防御式：所有外部输入(参数/配置/文件内容/SMTP 应答)都做校验与错误包装，不 panic；
//   * 维护性：按“配置/读取/摘要/发信/主流程”分段组织，每段职责单一。
//
// 构建：go build -o mailer(.exe) .
// 运行：./mailer -c config.json -f 新人磨合期记录.xlsx --dry-run   （预览自动摘要）
//       ./mailer -x -f 新人磨合期记录.docx                       （导出文本给 AI）
//       ./mailer -c config.json -s "新人磨合期记录·汇总" -b "……已总结正文……"（AI 模式直发）

package main

import (
	"archive/zip"
	"bytes"
	"crypto/rand"
	"crypto/tls"
	"encoding/base64"
	"encoding/csv"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"errors"
	"flag"
	"fmt"
	"io"
	"net"
	"net/smtp"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

// ---------------------------------------------------------------------------
// 常量与全局默认值
// ---------------------------------------------------------------------------

const (
	appName    = "本地文件内容汇报邮件发送器"
	appVersion = "1.2.0"

	defaultConfigName = "config.json" // -c 缺省时的配置文件
	// 主题规则：-s 参数 > config.defaults.subject(配置了才用) > 自动派生(正文/文件名概括) > 本兜底
	defaultSubject   = "内容汇报"
	defaultRecipient = "" // 发布版默认空：收件人须由配置 defaults.to 或 -t 提供（不在公开代码里内置收件人）

	envConfig   = "MAILER_CONFIG"   // 环境变量：配置文件路径
	envPassword = "MAILER_PASSWORD" // 环境变量：SMTP 授权码（推荐用于保护密钥，优先级高于配置文件）

	maxSourceBytes = int64(20 * 1024 * 1024) // 源文件体积上限 20MB
	maxBodyBytes   = int64(500 * 1024)       // 正文/正文文件体积上限 500KB
	defaultTimeout = 20                      // SMTP 连接超时（秒）

	emailMime = "dailymailer " + appVersion
)

// ---------------------------------------------------------------------------
// 通用小工具
// ---------------------------------------------------------------------------

// isPlaceholder 判断某配置项是否仍是“未填写”的占位符。
func isPlaceholder(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return true
	}
	markers := []string{"REPLACE", "xxxx", "XXXX", "YOUR_", "your_", "<your", "待填写", "请填写", "占位"}
	for _, m := range markers {
		if strings.Contains(s, m) {
			return true
		}
	}
	return false
}

func truncate(s string, maxRunes int) string {
	if utf8.RuneCountInString(s) <= maxRunes {
		return s
	}
	r := []rune(s)
	return string(r[:maxRunes]) + "…"
}

// trimSpaceCells 去除二维表格中每个单元格两侧空白。
func trimSpaceCells(rows [][]string) [][]string {
	out := make([][]string, 0, len(rows))
	for _, row := range rows {
		nr := make([]string, len(row))
		empty := true
		for i, c := range row {
			nr[i] = strings.TrimSpace(c)
			if nr[i] != "" {
				empty = false
			}
		}
		if !empty {
			out = append(out, nr)
		}
	}
	return out
}

// ---------------------------------------------------------------------------
// 邮件地址与 RFC2047 头部编码
// ---------------------------------------------------------------------------

type mailAddress struct {
	Name string // 显示名，可为空
	Addr string // 邮箱地址
}

// parseMailAddress 解析 “显示名 <a@b.c>” 或裸地址 “a@b.c”。
func parseMailAddress(raw string) (mailAddress, error) {
	s := strings.TrimSpace(raw)
	if s == "" {
		return mailAddress{}, errors.New("邮件地址为空")
	}
	if i := strings.LastIndexByte(s, '<'); i >= 0 {
		if j := strings.LastIndexByte(s, '>'); j > i {
			name := strings.TrimSpace(s[:i])
			addr := strings.TrimSpace(s[i+1 : j])
			if addr == "" {
				return mailAddress{}, fmt.Errorf("邮件地址缺少尖括号内的地址部分: %q", raw)
			}
			return mailAddress{Name: name, Addr: addr}, nil
		}
		return mailAddress{}, fmt.Errorf("邮件地址格式错误(缺少 >): %q", raw)
	}
	return mailAddress{Addr: s}, nil
}

func (a mailAddress) validate() error {
	if a.Addr == "" {
		return errors.New("邮件地址为空")
	}
	if strings.ContainsAny(a.Addr, " \t\r\n") {
		return fmt.Errorf("邮件地址含空白字符: %q", a.Addr)
	}
	if strings.Count(a.Addr, "@") != 1 {
		return fmt.Errorf("邮件地址必须且只能含一个 @: %q", a.Addr)
	}
	return nil
}

// headerString 生成用于 From/To/Cc 头部的字符串，非 ASCII 显示名做 RFC2047 编码。
func (a mailAddress) headerString() string {
	if a.Name != "" {
		return encodeRFC2047(a.Name) + " <" + a.Addr + ">"
	}
	return a.Addr
}

// encodeRFC2047 对非 ASCII 文本做 =?UTF-8?B?...?= 编码；纯 ASCII 原样返回。
func encodeRFC2047(s string) string {
	ascii := true
	for _, r := range s {
		if r > 127 {
			ascii = false
			break
		}
	}
	if ascii {
		return s
	}
	return "=?UTF-8?B?" + base64.StdEncoding.EncodeToString([]byte(s)) + "?="
}

// parseMailAddressList 批量解析并校验地址列表。
func parseMailAddressList(list []string) ([]mailAddress, error) {
	var out []mailAddress
	seen := map[string]bool{}
	for _, raw := range list {
		a, err := parseMailAddress(raw)
		if err != nil {
			return nil, err
		}
		if err := a.validate(); err != nil {
			return nil, fmt.Errorf("非法地址 %q: %w", raw, err)
		}
		key := strings.ToLower(a.Addr)
		if !seen[key] {
			seen[key] = true
			out = append(out, a)
		}
	}
	return out, nil
}

// ---------------------------------------------------------------------------
// 配置
// ---------------------------------------------------------------------------

type smtpConfig struct {
	Host               string `json:"host"`
	Port               int    `json:"port"`
	Username           string `json:"username"`
	Password           string `json:"password"`
	From               string `json:"from"`
	FromName           string `json:"fromName"`
	Security           string `json:"security"` // ssl | starttls | none；留空按端口推断
	InsecureSkipVerify bool   `json:"insecureSkipVerify"`
}

type defaultsConfig struct {
	To      []string `json:"to"`
	Cc      []string `json:"cc"`
	Subject string   `json:"subject"`
}

type config struct {
	Smtp           smtpConfig     `json:"smtp"`
	Defaults       defaultsConfig `json:"defaults"`
	TimeoutSeconds int            `json:"timeoutSeconds"`
}

// applyDefaults 填充缺省值（收件人/主题/超时/端口/安全模式等固定业务默认）。
func (c *config) applyDefaults() {
	if c.TimeoutSeconds <= 0 {
		c.TimeoutSeconds = defaultTimeout
	}
	if c.Smtp.Port == 0 {
		c.Smtp.Port = 465
	}
	if c.Smtp.Security == "" {
		switch c.Smtp.Port {
		case 465:
			c.Smtp.Security = "ssl"
		case 25, 587:
			c.Smtp.Security = "starttls"
		default:
			c.Smtp.Security = "ssl"
		}
	}
	if len(c.Defaults.To) == 0 && defaultRecipient != "" {
		c.Defaults.To = []string{defaultRecipient}
	}
	// 注意：defaults.subject 保持原样(可为空)。为空表示“主题自动概括文件内容”(见 run 中的主题派生)；
	//       若配置了非空值，则作为未传 -s 时的固定主题。
	if c.Smtp.FromName == "" {
		c.Smtp.FromName = "新人量化统计"
	}
}

// validateConfig 校验配置。skipCreds=true 时跳过账号/授权码类校验（用于 --dry-run 预览）。
func (c *config) validateConfig(skipCreds bool) error {
	var problems []string

	if strings.TrimSpace(c.Smtp.Host) == "" {
		problems = append(problems, "smtp.host 未填写")
	}
	if c.Smtp.Port < 1 || c.Smtp.Port > 65535 {
		problems = append(problems, fmt.Sprintf("smtp.port 非法: %d", c.Smtp.Port))
	}
	switch c.Smtp.Security {
	case "ssl", "starttls", "none":
	default:
		problems = append(problems, fmt.Sprintf("smtp.security 只能为 ssl/starttls/none，当前: %q", c.Smtp.Security))
	}

	if !skipCreds {
		if isPlaceholder(c.Smtp.Username) {
			problems = append(problems, "smtp.username(发件账号)未填写")
		}
		if isPlaceholder(c.Smtp.Password) {
			problems = append(problems, "smtp.password(授权码/密码)未填写，或仍为占位符")
		}
		if isPlaceholder(c.Smtp.From) {
			problems = append(problems, "smtp.from(发件人地址)未填写")
		}
	}
	// 说明：--dry-run(skipCreds) 允许账号/授权码仍为占位符，仅用于本地预览；
	//       from 的“格式”合法性仍由下方统一校验。

	if strings.TrimSpace(c.Smtp.From) != "" {
		if _, err := parseMailAddress(c.Smtp.From); err != nil {
			problems = append(problems, fmt.Sprintf("smtp.from 格式非法: %v", err))
		}
	}
	if _, err := parseMailAddressList(c.Defaults.To); err != nil {
		problems = append(problems, fmt.Sprintf("defaults.to 非法: %v", err))
	}
	if len(c.Defaults.Cc) > 0 {
		if _, err := parseMailAddressList(c.Defaults.Cc); err != nil {
			problems = append(problems, fmt.Sprintf("defaults.cc 非法: %v", err))
		}
	}

	if len(problems) > 0 {
		return fmt.Errorf("配置文件校验未通过：\n  - %s", strings.Join(problems, "\n  - "))
	}
	return nil
}

// loadConfig 读取并校验 JSON 配置文件。
// 说明：环境变量 MAILER_PASSWORD 在此生效——先于凭据校验，因此配置文件里 password 可留占位符，
//
//	仅通过环境变量提供授权码（推荐做法，避免密钥落盘）。
func loadConfig(path string, skipCreds bool) (*config, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, fmt.Errorf("找不到配置文件 %q：请先复制 config.example.json 为 config.json 并填写 SMTP 账号/授权码（或用 -c 显式指定路径）", path)
		}
		return nil, fmt.Errorf("读取配置文件 %q 失败: %w", path, err)
	}
	// 兼容带 BOM 的文件
	data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})

	var cfg config
	if err := json.Unmarshal(data, &cfg); err != nil {
		return nil, fmt.Errorf("解析配置文件 %q 失败(请检查 JSON 语法，如逗号/引号是否完整): %w", path, err)
	}
	cfg.applyDefaults()
	if pwd := os.Getenv(envPassword); pwd != "" {
		cfg.Smtp.Password = pwd
	}
	if err := cfg.validateConfig(skipCreds); err != nil {
		return nil, err
	}
	return &cfg, nil
}

// resolveConfigPath 依次取：-c 参数 > 环境变量 MAILER_CONFIG > 当前目录 config.json。
func resolveConfigPath(argPath string) string {
	if argPath != "" {
		return argPath
	}
	if p := os.Getenv(envConfig); p != "" {
		return p
	}
	return defaultConfigName
}

// ---------------------------------------------------------------------------
// 文档读取（table: xlsx/csv/tsv；text: txt/md/log 等）
// ---------------------------------------------------------------------------

type document struct {
	kind      string     // "table" 或 "text"
	table     [][]string // 表格模式：全部行
	textLines []string   // 文本模式：非空行
}

func newTableDoc(rows [][]string) *document {
	return &document{kind: "table", table: trimSpaceCells(rows)}
}

// readSourceFile 依据扩展名读取源文件（.xlsx/.csv/.tsv/.txt/.md/.log/.text）。
func readSourceFile(path string) (*document, error) {
	fi, err := os.Stat(path)
	if err != nil {
		return nil, fmt.Errorf("无法访问源文件 %q: %w", path, err)
	}
	if fi.IsDir() {
		return nil, fmt.Errorf("源路径 %q 是目录而非文件", path)
	}
	if fi.Size() > maxSourceBytes {
		return nil, fmt.Errorf("源文件 %q 超过体积上限 %.0fMB", path, float64(maxSourceBytes)/(1024*1024))
	}

	ext := strings.ToLower(filepath.Ext(path))
	switch ext {
	case ".xlsx":
		rows, err := readXLSX(path)
		if err != nil {
			return nil, fmt.Errorf("解析 xlsx %q 失败: %w", path, err)
		}
		return newTableDoc(rows), nil
	case ".csv", ".tsv":
		var delim rune = ','
		if ext == ".tsv" {
			delim = '\t'
		} else {
			delim = detectDelimiter(path)
		}
		rows, err := readDelimited(path, delim)
		if err != nil {
			return nil, fmt.Errorf("解析 %s %q 失败: %w", strings.TrimPrefix(ext, "."), path, err)
		}
		return newTableDoc(rows), nil
	case ".txt", ".md", ".log", ".text", ".markdown":
		lines, err := readTextLines(path)
		if err != nil {
			return nil, fmt.Errorf("读取文本 %q 失败: %w", path, err)
		}
		return &document{kind: "text", textLines: lines}, nil
	case ".docx":
		lines, err := parseDOCX(path)
		if err != nil {
			return nil, err
		}
		return &document{kind: "text", textLines: lines}, nil
	case ".pptx":
		lines, err := parsePPTX(path)
		if err != nil {
			return nil, err
		}
		return &document{kind: "text", textLines: lines}, nil
	default:
		return nil, fmt.Errorf("不支持的源文件类型 %q（支持表格 .xlsx/.csv/.tsv，文档 .docx/.pptx，文本 .txt/.md/.log/.text）", ext)
	}
}

// detectDelimiter 对 .csv 做启发式分隔符探测（防止“分号分隔”的 Excel 导出被误读）。
func detectDelimiter(path string) rune {
	f, err := os.Open(path)
	if err != nil {
		return ','
	}
	defer f.Close()
	buf := make([]byte, 8192)
	n, _ := f.Read(buf)
	if n == 0 {
		return ','
	}
	head := string(buf[:n])
	if i := strings.IndexAny(head, "\r\n"); i >= 0 {
		head = head[:i]
	}
	count := func(r rune) int { return strings.Count(head, string(r)) }
	if count(';') > count(',') && count(';') > count('\t') {
		return ';'
	}
	if count('\t') > count(',') && count('\t') > count(';') {
		return '\t'
	}
	return ','
}

// readDelimited 读取分隔符文本(csv/tsv)。
func readDelimited(path string, delim rune) ([][]string, error) {
	data, err := readFileBytes(path)
	if err != nil {
		return nil, err
	}
	if !utf8.Valid(data) {
		return nil, errors.New("文件不是合法 UTF-8 编码（常见于 GBK/ANSI 导出）。请在 Excel/记事本中另存为“UTF-8”后重试，或改从腾讯文档导出 xlsx")
	}
	data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})
	r := csv.NewReader(bytes.NewReader(data))
	r.Comma = delim
	r.FieldsPerRecord = -1 // 每行列数可变，宽松处理
	r.LazyQuotes = true    // 容忍不规范引号
	rows, err := r.ReadAll()
	if err != nil {
		return nil, fmt.Errorf("无法按分隔符解析(可能文件不是真正以 %q 分隔): %w", string(delim), err)
	}
	return trimSpaceCells(rows), nil
}

// readTextLines 读取纯文本，返回非空行。
func readTextLines(path string) ([]string, error) {
	data, err := readFileBytes(path)
	if err != nil {
		return nil, err
	}
	if !utf8.Valid(data) {
		return nil, errors.New("文件不是合法 UTF-8 编码。请在编辑器中转存为 UTF-8 后重试")
	}
	text := string(bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF}))
	var lines []string
	for _, ln := range strings.Split(text, "\n") {
		ln = strings.TrimRight(strings.TrimSpace(ln), "\r")
		if ln != "" {
			lines = append(lines, ln)
		}
	}
	return lines, nil
}

func readFileBytes(path string) ([]byte, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	data, err := io.ReadAll(io.LimitReader(f, maxSourceBytes+1))
	if err != nil {
		return nil, err
	}
	if int64(len(data)) > maxSourceBytes {
		return nil, fmt.Errorf("文件超过体积上限 %.0fMB", float64(maxSourceBytes)/(1024*1024))
	}
	return data, nil
}

// ---------------------------------------------------------------------------
// xlsx 最小解析器（仅标准库；覆盖腾讯文档/Excel 常见导出的字符串与数值单元格）
// ---------------------------------------------------------------------------

type xlsxSst struct {
	Si []struct {
		T    []string `xml:"t"` // 简单字符串的 <t>
		Runs []struct {
			T []string `xml:"t"` // 富文本 run 内的 <t>
		} `xml:"r"`
	} `xml:"si"`
}

type xlsxCell struct {
	Ref  string `xml:"r,attr"`
	Type string `xml:"t,attr"`
	V    string `xml:"v"`
	Is   struct {
		T []string `xml:"t"`
	} `xml:"is"`
}

type xlsxWorksheet struct {
	SheetData struct {
		Rows []struct {
			Cells []xlsxCell `xml:"c"`
		} `xml:"row"`
	} `xml:"sheetData"`
}

type xlsxWorkbook struct {
	Sheets []struct {
		RID string `xml:"id,attr"` // 即 r:id 的本地名 id
	} `xml:"sheets>sheet"`
}

type xlsxRels struct {
	Relationships []struct {
		ID     string `xml:"Id,attr"`
		Target string `xml:"Target,attr"`
	} `xml:"Relationship"`
}

// columnIndexOfRef 把 “AB12” 的列字母部分转成 0 起始列下标。
func columnIndexOfRef(ref string) (int, bool) {
	idx := 0
	seen := false
	for _, b := range ref {
		if b >= 'A' && b <= 'Z' {
			idx = idx*26 + int(b-'A'+1)
			seen = true
		} else if b >= 'a' && b <= 'z' {
			idx = idx*26 + int(b-'a'+1)
			seen = true
		} else {
			break
		}
	}
	if !seen {
		return 0, false
	}
	return idx - 1, true
}

// readXLSX 打开 .xlsx(zip) 并读取第一个工作表的所有行。
func readXLSX(path string) ([][]string, error) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("无法打开 xlsx(zip) 包: %w", err)
	}
	defer zr.Close()

	byName := map[string]*zip.File{}
	for _, f := range zr.File {
		byName[strings.ToLower(f.Name)] = f
	}
	find := func(suffix string) *zip.File {
		if f, ok := byName[suffix]; ok {
			return f
		}
		var cands []*zip.File
		for _, f := range zr.File {
			if strings.HasSuffix(strings.ToLower(f.Name), suffix) {
				cands = append(cands, f)
			}
		}
		if len(cands) == 0 {
			return nil
		}
		sort.Slice(cands, func(i, j int) bool { return cands[i].Name < cands[j].Name })
		return cands[0]
	}

	// 1) 共享字符串表
	var shared []string
	if sstFile := find("/sharedstrings.xml"); sstFile != nil {
		data, err := readZipEntry(sstFile)
		if err != nil {
			return nil, err
		}
		var sst xlsxSst
		if err := xml.Unmarshal(data, &sst); err == nil {
			for _, si := range sst.Si {
				var sb strings.Builder
				for _, t := range si.T {
					sb.WriteString(t)
				}
				for _, run := range si.Runs {
					for _, t := range run.T {
						sb.WriteString(t)
					}
				}
				shared = append(shared, sb.String())
			}
		}
	}

	// 2) 依据 workbook 顺序找第一个工作表
	sheetRel := find("/workbook.xml")
	var sheetPath string
	if sheetRel != nil {
		if data, err := readZipEntry(sheetRel); err == nil {
			var wb xlsxWorkbook
			if xml.Unmarshal(data, &wb) == nil && len(wb.Sheets) > 0 {
				rid := wb.Sheets[0].RID
				if relsFile := find("/workbook.xml.rels"); relsFile != nil {
					if rd, err := readZipEntry(relsFile); err == nil {
						var rels xlsxRels
						if xml.Unmarshal(rd, &rels) == nil {
							for _, r := range rels.Relationships {
								if r.ID == rid && r.Target != "" {
									t := strings.TrimPrefix(strings.ReplaceAll(r.Target, "\\", "/"), "/")
									if !strings.HasPrefix(strings.ToLower(t), "xl/") {
										t = "xl/" + t
									}
									sheetPath = t
									break
								}
							}
						}
					}
				}
			}
		}
	}
	if sheetPath == "" {
		// 兜底：取名称排序后第一个 xl/worksheets/sheet*.xml
		var cands []*zip.File
		for _, f := range zr.File {
			n := strings.ToLower(f.Name)
			if strings.HasPrefix(n, "xl/worksheets/") && strings.HasSuffix(n, ".xml") && !strings.Contains(n, "rels") {
				cands = append(cands, f)
			}
		}
		if len(cands) == 0 {
			return nil, errors.New("xlsx 中找不到工作表(xl/worksheets/*.xml)")
		}
		sort.Slice(cands, func(i, j int) bool { return cands[i].Name < cands[j].Name })
		sheetPath = cands[0].Name
	}

	wsFile := byName[strings.ToLower(sheetPath)]
	if wsFile == nil {
		wsFile = find("/" + strings.ToLower(filepath.Base(sheetPath)))
	}
	if wsFile == nil {
		return nil, fmt.Errorf("无法定位工作表 %q", sheetPath)
	}
	data, err := readZipEntry(wsFile)
	if err != nil {
		return nil, err
	}
	var ws xlsxWorksheet
	if err := xml.Unmarshal(data, &ws); err != nil {
		return nil, fmt.Errorf("解析工作表 XML 失败: %w", err)
	}

	// 3) 组装行
	var rows [][]string
	for _, r := range ws.SheetData.Rows {
		cells := map[int]string{}
		maxCol := -1
		next := 0
		for _, c := range r.Cells {
			col, ok := columnIndexOfRef(c.Ref)
			if !ok {
				col = next
			}
			next = col + 1
			v := xlsxCellValue(c, shared)
			if v != "" {
				if col > maxCol {
					maxCol = col
				}
				cells[col] = v
			}
		}
		if maxCol < 0 {
			continue // 整行为空
		}
		row := make([]string, maxCol+1)
		for i := 0; i <= maxCol; i++ {
			row[i] = cells[i]
		}
		rows = append(rows, row)
	}
	if len(rows) == 0 {
		return nil, errors.New("xlsx 第一个工作表没有读到任何数据行")
	}
	return trimSpaceCells(rows), nil
}

// xlsxCellValue 依据单元格类型取显示值。
func xlsxCellValue(c xlsxCell, shared []string) string {
	switch c.Type {
	case "s": // shared string
		i, err := strconv.Atoi(strings.TrimSpace(c.V))
		if err != nil || i < 0 || i >= len(shared) {
			return ""
		}
		return shared[i]
	case "inlineStr": // 内联字符串
		return strings.Join(c.Is.T, "")
	case "str": // 公式字符串结果
		return c.V
	case "b": // 布尔
		if c.V == "1" {
			return "是"
		}
		return "否"
	default: // 数值、日期序列号等，一律按文本保留
		return strings.TrimSpace(c.V)
	}
}

func readZipEntry(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(io.LimitReader(rc, maxSourceBytes+1))
}

// ---------------------------------------------------------------------------
// Office 文档（docx/pptx）文本抽取：本质是 zip + XML，仅用标准库
// ---------------------------------------------------------------------------

// extractParagraphLines 从 OOXML XML 中按 <p> 段落抽取文本行
// （按元素 local name 匹配，兼容 w:/a: 等命名空间前缀；段落内所有 <t> 文本会被拼接）。
func extractParagraphLines(data []byte) ([]string, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	paraDepth := 0
	var cur strings.Builder
	var lines []string
	var stack []string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			ln := t.Name.Local
			stack = append(stack, ln)
			if ln == "p" {
				paraDepth++
			}
		case xml.EndElement:
			ln := t.Name.Local
			if ln == "p" {
				paraDepth--
				if s := strings.TrimSpace(cur.String()); s != "" {
					lines = append(lines, s)
				}
				cur.Reset()
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			if paraDepth > 0 && len(stack) > 0 && stack[len(stack)-1] == "t" {
				cur.Write(t)
			}
		}
	}
	return lines, nil
}

// zipEntryBySuffix 返回压缩包内文件名(小写)以 suffix 结尾的首个条目。
func zipEntryBySuffix(zr *zip.ReadCloser, suffix string) (*zip.File, bool) {
	for _, f := range zr.File {
		if strings.HasSuffix(strings.ToLower(f.Name), suffix) {
			return f, true
		}
	}
	return nil, false
}

// parseDOCX 抽取 Word 正文（word/document.xml）的段落文本。
func parseDOCX(path string) ([]string, error) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("无法打开 docx(zip) 包: %w", err)
	}
	defer zr.Close()
	f, ok := zipEntryBySuffix(zr, "/document.xml")
	if !ok {
		return nil, errors.New("docx 中找不到 word/document.xml")
	}
	data, err := readZipEntry(f)
	if err != nil {
		return nil, err
	}
	lines, err := extractParagraphLines(data)
	if err != nil {
		return nil, fmt.Errorf("解析 document.xml 失败: %w", err)
	}
	return lines, nil
}

// parsePPTX 按页顺序抽取幻灯片文本，每页前插入“【第N页】”标记。
func parsePPTX(path string) ([]string, error) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("无法打开 pptx(zip) 包: %w", err)
	}
	defer zr.Close()
	type slideEntry struct {
		idx  int
		file *zip.File
	}
	var slides []slideEntry
	for _, f := range zr.File {
		n := strings.ToLower(f.Name)
		if !strings.HasPrefix(n, "ppt/slides/slide") || !strings.HasSuffix(n, ".xml") || strings.Contains(n, "rels") {
			continue
		}
		base := filepath.Base(f.Name) // slideN.xml
		num := strings.TrimSuffix(strings.TrimPrefix(strings.ToLower(base), "slide"), ".xml")
		idx, err := strconv.Atoi(num)
		if err != nil {
			continue
		}
		slides = append(slides, slideEntry{idx: idx, file: f})
	}
	if len(slides) == 0 {
		return nil, errors.New("pptx 中找不到 ppt/slides/slide*.xml")
	}
	sort.Slice(slides, func(i, j int) bool { return slides[i].idx < slides[j].idx })
	var out []string
	for _, s := range slides {
		data, err := readZipEntry(s.file)
		if err != nil {
			return nil, err
		}
		lines, err := extractParagraphLines(data)
		if err != nil {
			return nil, fmt.Errorf("解析幻灯片 %d 失败: %w", s.idx, err)
		}
		out = append(out, fmt.Sprintf("【第%d页】", s.idx))
		out = append(out, lines...)
	}
	return out, nil
}

// ---------------------------------------------------------------------------
// 文本/表格的“可读文本”导出（供 -x/--extract 输出给 AI Agent 阅读）
// ---------------------------------------------------------------------------

func rowLooksLikeNote(row []string) bool {
	total := 0
	for _, c := range row {
		if c == "" {
			continue
		}
		if utf8.RuneCountInString(c) > 200 {
			return true
		}
		total += utf8.RuneCountInString(c)
	}
	return total > 500
}

func rowNonEmptyCount(row []string) int {
	n := 0
	for _, c := range row {
		if c != "" {
			n++
		}
	}
	return n
}

// tableExtractText 把表格转成“表头行+数据行”的制表符分隔文本，便于 AI 阅读/总结。
func tableExtractText(rows [][]string) string {
	// 与 tableAutoSummary 一致的表头探测：跳过标题/说明行，取首个“非说明且非空列数达标”的行当表头
	maxNE := 0
	for _, r := range rows {
		if v := rowNonEmptyCount(r); v > maxNE {
			maxNE = v
		}
	}
	threshold := maxNE * 3 / 5
	if threshold < 2 {
		threshold = 2
	}
	headerIdx := -1
	for i, r := range rows {
		if !rowLooksLikeNote(r) && rowNonEmptyCount(r) >= threshold {
			headerIdx = i
			break
		}
	}
	var header, body [][]string
	if len(rows) == 1 {
		body = rows
	} else if headerIdx >= 0 {
		header = rows[headerIdx : headerIdx+1]
		for _, r := range rows[headerIdx+1:] {
			if !rowLooksLikeNote(r) {
				body = append(body, r)
			}
		}
	} else {
		for _, r := range rows {
			if !rowLooksLikeNote(r) {
				body = append(body, r)
			}
		}
	}
	colName := func(i int) string {
		if header != nil && i < len(header[0]) && header[0][i] != "" {
			return header[0][i]
		}
		return fmt.Sprintf("第%d列", i+1)
	}
	cell := func(i int, c string) string {
		if isDateLikeHeader(colName(i)) {
			if d, ok := excelSerialToDate(c); ok {
				return d
			}
		}
		return c
	}
	ncols := 0
	for _, r := range body {
		if len(r) > ncols {
			ncols = len(r)
		}
	}
	var sb strings.Builder
	if header != nil {
		var hs []string
		for i := 0; i < ncols; i++ {
			hs = append(hs, header[0][i])
		}
		sb.WriteString(strings.Join(hs, "\t"))
		sb.WriteString("\n")
	}
	for _, row := range body {
		var cells []string
		for i := 0; i < ncols; i++ {
			c := ""
			if i < len(row) {
				c = row[i]
			}
			cells = append(cells, truncate(cell(i, c), 300))
		}
		sb.WriteString(strings.Join(cells, "\t"))
		sb.WriteString("\n")
	}
	return sb.String()
}

// docExtractText 输出文件的可读文本（-x 用；表格输出制表符分隔，Office/纯文本按行原样输出）。
func docExtractText(doc *document) string {
	if doc.kind == "table" {
		return tableExtractText(doc.table)
	}
	var sb strings.Builder
	for _, l := range doc.textLines {
		sb.WriteString(l)
		sb.WriteString("\n")
	}
	return sb.String()
}

func firstNonEmptyLine(s string) string {
	for _, ln := range strings.Split(s, "\n") {
		if t := strings.TrimSpace(ln); t != "" {
			return t
		}
	}
	return ""
}

// ---------------------------------------------------------------------------
// 自动摘要（仅作为“无 -b 正文”时的确定性兜底，不替代 AI 精炼）
// ---------------------------------------------------------------------------

func isDateLikeHeader(s string) bool {
	kw := []string{"日期", "时间", "date", "time", "day", "周", "月份"}
	ls := strings.ToLower(s)
	for _, k := range kw {
		if strings.Contains(ls, strings.ToLower(k)) {
			return true
		}
	}
	return false
}

// isRateLikeHeader 判断列名是否为“比率/百分比”语义（如完成率、达标率、占比）。
func isRateLikeHeader(s string) bool {
	kw := []string{"率", "%", "％", "占比", "percentage"}
	ls := strings.ToLower(s)
	for _, k := range kw {
		if strings.Contains(ls, strings.ToLower(k)) {
			return true
		}
	}
	return false
}

// excelSerialToDate 把 Excel 日期序列号(自 1899-12-30 起的天数)转成 yyyy-MM-dd。
func excelSerialToDate(s string) (string, bool) {
	v, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil || v < 20000 || v > 80000 {
		return "", false
	}
	day := int(v)
	if v-float64(day) >= 0.5 { // 处理带时间的序列号，四舍五入到天
		day++
	}
	d := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).AddDate(0, 0, day)
	return d.Format("2006-01-02"), true
}

func parseNumberCell(s string) (float64, bool) {
	t := strings.TrimSpace(s)
	t = strings.ReplaceAll(t, ",", "")
	t = strings.ReplaceAll(t, "\u00a0", "")
	t = strings.Trim(t, " ￥¥$€")
	if t == "" || strings.ContainsAny(t, "%％") {
		return 0, false
	}
	v, err := strconv.ParseFloat(t, 64)
	if err != nil {
		return 0, false
	}
	return v, true
}

func fmtNum(v float64) string {
	if v == float64(int64(v)) {
		return strconv.FormatInt(int64(v), 10)
	}
	return strconv.FormatFloat(v, 'f', 2, 64)
}

// tableAutoSummary 生成表格的确定性摘要（中文、一行行组成的一段话）。
func tableAutoSummary(rows [][]string) string {
	var sb strings.Builder

	// —— 表头识别（兼容“标题/说明块 + 表头行 + 数据行”等混合布局）——
	// 说明/注解类行：存在超长单元格(>200字)或整行文本合计很长(>500字)
	isNoteRow := func(row []string) bool {
		total := 0
		for _, c := range row {
			if c == "" {
				continue
			}
			if utf8.RuneCountInString(c) > 200 {
				return true
			}
			total += utf8.RuneCountInString(c)
		}
		return total > 500
	}
	nonEmpty := func(row []string) int {
		n := 0
		for _, c := range row {
			if c != "" {
				n++
			}
		}
		return n
	}
	maxNE := 0
	for _, r := range rows {
		if v := nonEmpty(r); v > maxNE {
			maxNE = v
		}
	}
	// 表头候选：非空列数 ≥ 全表最大非空列数的 60% 的第一条“非说明”行
	threshold := maxNE * 3 / 5
	if threshold < 2 {
		threshold = 2
	}
	headerIdx := -1
	for i, r := range rows {
		if !isNoteRow(r) && nonEmpty(r) >= threshold {
			headerIdx = i
			break
		}
	}

	var header, body [][]string
	hasHeader := false
	if len(rows) == 1 {
		body = rows // 仅一行：按纯数据行处理，不设表头
	} else if headerIdx >= 0 {
		hasHeader = true
		header = rows[headerIdx : headerIdx+1]
		for _, r := range rows[headerIdx+1:] {
			if !isNoteRow(r) {
				body = append(body, r)
			}
		}
	} else {
		for _, r := range rows {
			if !isNoteRow(r) {
				body = append(body, r)
			}
		}
	}
	skippedLead := 0
	if hasHeader {
		skippedLead = headerIdx
	} else {
		for i, r := range rows {
			if !isNoteRow(r) {
				skippedLead = i
				break
			}
		}
	}
	headerRows := 0
	if hasHeader {
		headerRows = 1
	}
	skippedTotal := len(rows) - len(body) - headerRows

	ncols := 0
	for _, r := range body {
		if len(r) > ncols {
			ncols = len(r)
		}
	}
	colName := func(i int) string {
		if hasHeader && i < len(header[0]) && header[0][i] != "" {
			return header[0][i]
		}
		return fmt.Sprintf("第%d列", i+1)
	}
	// display 用于展示层：日期列若为 Excel 序列号则显示为真实日期
	display := func(i int, c string) string {
		if isDateLikeHeader(colName(i)) {
			if d, ok := excelSerialToDate(c); ok {
				return d
			}
		}
		return c
	}

	fmt.Fprintf(&sb, "共 %d 条数据记录、%d 列", len(body), ncols)
	if hasHeader {
		var hs []string
		for i := 0; i < ncols; i++ {
			if i < len(header[0]) {
				hs = append(hs, header[0][i])
			} else {
				hs = append(hs, fmt.Sprintf("第%d列", i+1))
			}
		}
		sb.WriteString("，表头：" + strings.Join(hs, " | "))
	}
	sb.WriteString("。")
	if skippedTotal > 0 {
		var parts []string
		if skippedLead > 0 {
			parts = append(parts, fmt.Sprintf("开头跳过 %d 行标题/说明", skippedLead))
		}
		if skippedTotal > skippedLead {
			parts = append(parts, fmt.Sprintf("另有 %d 行说明/注解行未计入", skippedTotal-skippedLead))
		}
		sb.WriteString("\n（已忽略：" + strings.Join(parts, "；") + "）")
	}

	// 各列统计
	type colStat struct {
		nonEmpty int
		sum      float64
		sumN     int
		minV     float64
		maxV     float64
		haveNum  bool
		samples  []string
		allNum   bool
		isRate   bool
		any      bool
	}
	stats := make([]colStat, ncols)
	for _, row := range body {
		for i := 0; i < ncols; i++ {
			c := ""
			if i < len(row) {
				c = row[i]
			}
			if c == "" {
				continue
			}
			st := &stats[i]
			st.any = true
			st.nonEmpty++
			if len(st.samples) < 3 {
				st.samples = append(st.samples, truncate(display(i, c), 20))
			}
			if v, ok := parseNumberCell(c); ok {
				st.sum += v
				st.sumN++
				if !st.haveNum {
					st.haveNum = true
					st.minV = v
					st.maxV = v
				} else {
					if v < st.minV {
						st.minV = v
					}
					if v > st.maxV {
						st.maxV = v
					}
				}
			}
		}
	}
	for i := range stats {
		st := &stats[i]
		if !st.any {
			continue
		}
		st.allNum = st.sumN > 0 && st.sumN == st.nonEmpty && !isDateLikeHeader(colName(i))
		// 比率/百分比列：全部数值 ∈ [0,1] 且列名含“率/%/占比”时按平均百分比展示，不再累加
		st.isRate = st.allNum && isRateLikeHeader(colName(i)) && st.minV >= 0 && st.maxV <= 1
	}

	// 数值列（最多展示 8 个，其余汇总提示）
	var numParts []string
	numShown := 0
	numTotal := 0
	for i := 0; i < ncols; i++ {
		if stats[i].any && stats[i].allNum {
			numTotal++
		}
	}
	for i := 0; i < ncols && numShown < 8; i++ {
		st := &stats[i]
		if !st.any || !st.allNum {
			continue
		}
		avg := st.sum / float64(st.sumN)
		if st.isRate {
			numParts = append(numParts, fmt.Sprintf("%s（平均 %s）", colName(i), fmtNum(avg*100)+"%"))
		} else {
			numParts = append(numParts, fmt.Sprintf("%s（合计%s，平均%s）", colName(i), fmtNum(st.sum), fmtNum(avg)))
		}
		numShown++
	}
	if len(numParts) > 0 {
		sb.WriteString("\n数值列统计：" + strings.Join(numParts, "；"))
		if numTotal > numShown {
			fmt.Fprintf(&sb, "；另有 %d 个数值列从略", numTotal-numShown)
		}
		sb.WriteString("。")
	}

	// 文本列取值示例（最多展示 4 个非空文本列）
	shown := 0
	for i := 0; i < ncols && shown < 4; i++ {
		if !stats[i].any || stats[i].allNum {
			continue
		}
		sb.WriteString(fmt.Sprintf("\n%s 取值如：%s。", colName(i), strings.Join(stats[i].samples, "、")))
		shown++
	}

	// 前 3 条记录预览
	if len(body) > 0 {
		sb.WriteString("\n记录预览（前 ")
		n := len(body)
		if n > 3 {
			n = 3
		}
		fmt.Fprintf(&sb, "%d 条）：", n)
		for ri := 0; ri < n; ri++ {
			var parts []string
			for i := 0; i < ncols; i++ {
				c := ""
				if i < len(body[ri]) {
					c = body[ri][i]
				}
				parts = append(parts, truncate(display(i, c), 24))
			}
			fmt.Fprintf(&sb, "%d) %s；", ri+1, strings.Join(parts, " / "))
		}
		sb.WriteString("。")
	}
	return sb.String()
}

func textAutoSummary(lines []string) string {
	total := 0
	for _, l := range lines {
		total += utf8.RuneCountInString(l)
	}
	var sb strings.Builder
	fmt.Fprintf(&sb, "该文件为纯文本，共 %d 行、约 %d 字。", len(lines), total)
	if len(lines) > 0 {
		n := len(lines)
		if n > 5 {
			n = 5
		}
		sb.WriteString("\n开头内容摘录：\n")
		for i := 0; i < n; i++ {
			sb.WriteString("  " + truncate(lines[i], 80) + "\n")
		}
		if len(lines) > 5 {
			fmt.Fprintf(&sb, "  ……（其余 %d 行省略）", len(lines)-5)
		}
	}
	return sb.String()
}

// composeAutoBody 组装“自动摘要模式”的邮件正文。
func composeAutoBody(doc *document, srcName string) string {
	var summary string
	if doc.kind == "table" {
		summary = tableAutoSummary(doc.table)
	} else {
		summary = textAutoSummary(doc.textLines)
	}
	return fmt.Sprintf(
		"文件《%s》内容自动汇总：\n\n%s\n\n（本段由脚本按文件结构自动汇总生成；如需更精炼自然的表述，"+
			"请让 AI 依据 AI skill 阅读文档后总结，再用 -b 参数传入正文。）",
		srcName, summary)
}

// ---------------------------------------------------------------------------
// SMTP 发信
// ---------------------------------------------------------------------------

// loginAuth 实现 SMTP AUTH LOGIN（部分服务器只提供 LOGIN）。
type loginAuth struct {
	username, password string
}

func (a loginAuth) Start(_ *smtp.ServerInfo) (string, []byte, error) {
	return "LOGIN", nil, nil
}

func (a loginAuth) Next(fromServer []byte, more bool) ([]byte, error) {
	if !more {
		return nil, nil
	}
	msg := strings.ToLower(string(fromServer))
	switch {
	case strings.HasPrefix(msg, "user"):
		return []byte(a.username), nil
	case strings.HasPrefix(msg, "pass"):
		return []byte(a.password), nil
	default:
		return nil, fmt.Errorf("LOGIN 认证收到未知服务器提示: %q", string(fromServer))
	}
}

// buildMessage 组装符合 RFC5322 的邮件原文（含 RFC2047 编码的中文主题/显示名、base64 正文）。
func buildMessage(from mailAddress, fromName string, to, cc []mailAddress, subject, body string) []byte {
	var hdrs []string
	add := func(k, v string) {
		hdrs = append(hdrs, k+": "+v)
	}
	sender := mailAddress{Name: fromName, Addr: from.Addr}
	add("From", sender.headerString())
	var tos []string
	for _, a := range to {
		tos = append(tos, a.headerString())
	}
	add("To", strings.Join(tos, ", "))
	if len(cc) > 0 {
		var ccs []string
		for _, a := range cc {
			ccs = append(ccs, a.headerString())
		}
		add("Cc", strings.Join(ccs, ", "))
	}
	add("Subject", encodeRFC2047(subject))
	add("Date", time.Now().Format(time.RFC1123Z))
	add("Message-ID", newMessageID(from.Addr))
	add("MIME-Version", "1.0")
	add("Content-Type", "text/plain; charset=UTF-8")
	add("Content-Transfer-Encoding", "base64")
	add("X-Mailer", emailMime)

	header := strings.Join(hdrs, "\r\n") + "\r\n\r\n"
	encoded := base64.StdEncoding.EncodeToString([]byte(body))
	var b bytes.Buffer
	b.WriteString(header)
	for len(encoded) > 76 {
		b.WriteString(encoded[:76])
		b.WriteString("\r\n")
		encoded = encoded[76:]
	}
	b.WriteString(encoded)
	b.WriteString("\r\n")
	return b.Bytes()
}

func newMessageID(host string) string {
	var raw [12]byte
	if _, err := rand.Read(raw[:]); err != nil {
		return fmt.Sprintf("<%d.%d@%s>", time.Now().UnixNano(), os.Getpid(), sanitizeHost(host))
	}
	return "<" + hex.EncodeToString(raw[:]) + "@" + sanitizeHost(host) + ">"
}

func sanitizeHost(h string) string {
	h = strings.TrimSpace(h)
	if h == "" {
		return "localhost"
	}
	if i := strings.LastIndexByte(h, ':'); i >= 0 {
		h = h[:i]
	}
	if strings.ContainsAny(h, "<> \t") {
		return "localhost"
	}
	return h
}

// sendMail 通过 SMTP 发送邮件。支持 ssl(465)/starttls(587,25)/none。
func sendMail(cfg *config, to, cc []mailAddress, subject, body string, verbose bool) error {
	smtpCfg := cfg.Smtp
	host := strings.TrimSpace(smtpCfg.Host)
	if host == "" {
		return errors.New("smtp.host 为空")
	}
	addr := net.JoinHostPort(host, strconv.Itoa(smtpCfg.Port))
	timeout := time.Duration(cfg.TimeoutSeconds) * time.Second

	vlog := func(format string, a ...interface{}) {
		if verbose {
			fmt.Fprintf(os.Stderr, "[mailer] "+format+"\n", a...)
		}
	}

	fromAddr, err := parseMailAddress(smtpCfg.From)
	if err != nil {
		return err
	}
	if err := fromAddr.validate(); err != nil {
		return fmt.Errorf("发件人地址非法: %w", err)
	}

	conn, err := net.DialTimeout("tcp", addr, timeout)
	if err != nil {
		return fmt.Errorf("连接 SMTP 服务器 %s 超时/失败: %w", addr, err)
	}
	defer conn.Close()
	vlog("已连接 %s", addr)

	tlsCfg := &tls.Config{
		ServerName:         host,
		InsecureSkipVerify: smtpCfg.InsecureSkipVerify,
		MinVersion:         tls.VersionTLS12,
	}

	var client *smtp.Client
	switch smtpCfg.Security {
	case "ssl":
		tconn := tls.Client(conn, tlsCfg)
		if err := tconn.Handshake(); err != nil {
			return fmt.Errorf("SMTP SSL 握手失败(确认端口与服务器支持 465 隐式 TLS): %w", err)
		}
		client, err = smtp.NewClient(tconn, host)
		if err != nil {
			return fmt.Errorf("SMTP 会话初始化失败: %w", err)
		}
	case "starttls":
		client, err = smtp.NewClient(conn, host)
		if err != nil {
			return fmt.Errorf("SMTP 会话初始化失败: %w", err)
		}
		if ok, _ := client.Extension("STARTTLS"); !ok {
			client.Close()
			return fmt.Errorf("SMTP 服务器 %s 不支持 STARTTLS", addr)
		}
		if err := client.StartTLS(tlsCfg); err != nil {
			client.Close()
			return fmt.Errorf("SMTP STARTTLS 升级失败: %w", err)
		}
	default: // none
		client, err = smtp.NewClient(conn, host)
		if err != nil {
			return fmt.Errorf("SMTP 会话初始化失败: %w", err)
		}
	}
	defer client.Close()
	vlog("SMTP 会话已建立（安全模式 %s）", smtpCfg.Security)

	// 认证：优先使用服务器公告的机制；LOGIN 优先，其次 PLAIN
	if strings.TrimSpace(smtpCfg.Username) != "" {
		mechs := ""
		if ok, m := client.Extension("AUTH"); ok {
			mechs = strings.ToUpper(m)
		}
		vlog("服务器 AUTH 机制: %s", strings.TrimSpace(mechs))
		var authErr error
		if strings.Contains(mechs, "LOGIN") {
			authErr = client.Auth(loginAuth{username: smtpCfg.Username, password: smtpCfg.Password})
		} else if strings.Contains(mechs, "PLAIN") {
			authErr = client.Auth(smtp.PlainAuth("", smtpCfg.Username, smtpCfg.Password, host))
		} else {
			return fmt.Errorf("SMTP 服务器未公告支持的认证机制(LOGIN/PLAIN)，请检查账号配置")
		}
		if authErr != nil {
			return fmt.Errorf("SMTP 认证失败(请检查发件账号与授权码是否正确、授权码是否过期): %w", authErr)
		}
		vlog("SMTP 认证成功: %s", smtpCfg.Username)
	} else {
		return errors.New("未配置 smtp.username 发件账号")
	}

	// 发信
	if err := client.Mail(fromAddr.Addr); err != nil {
		return fmt.Errorf("MAIL FROM 被拒绝: %w", err)
	}
	rcpts := make([]string, 0, len(to)+len(cc))
	for _, a := range append(append([]mailAddress{}, to...), cc...) {
		if err := client.Rcpt(a.Addr); err != nil {
			return fmt.Errorf("RCPT TO %s 被拒绝: %w", a.Addr, err)
		}
		rcpts = append(rcpts, a.Addr)
	}
	vlog("收件人: %s", strings.Join(rcpts, ", "))

	w, err := client.Data()
	if err != nil {
		return fmt.Errorf("DATA 命令失败: %w", err)
	}
	msg := buildMessage(fromAddr, smtpCfg.FromName, to, cc, subject, body)
	if _, err := w.Write(msg); err != nil {
		w.Close()
		return fmt.Errorf("写入邮件内容失败: %w", err)
	}
	if err := w.Close(); err != nil {
		return fmt.Errorf("发送邮件内容失败(连接可能中断): %w", err)
	}
	if err := client.Quit(); err != nil {
		// 部分服务器在 QUIT 时已断开；此处不视为致命错误
		vlog("QUIT 提示: %v", err)
	}
	return nil
}

// ---------------------------------------------------------------------------
// 命令行参数与帮助
// ---------------------------------------------------------------------------

type cliOptions struct {
	configPath string
	filePath   string
	bodyParts  multiFlag
	toList     addrList
	subject    string
	dryRun     bool
	extract    bool
	verbose    bool
	help       bool
}

// multiFlag 允许 -b 重复出现，自动以换行拼接（便于传入多段正文）。
type multiFlag []string

func (m *multiFlag) String() string { return strings.Join(*m, "\n") }
func (m *multiFlag) Set(v string) error {
	*m = append(*m, v)
	return nil
}

// addrList 允许 -t/--to 指定一个或多个收件人，支持重复传入或用 , ; ， ； 分隔。
type addrList []string

func (a *addrList) String() string { return strings.Join(*a, ",") }
func (a *addrList) Set(v string) error {
	for _, part := range strings.FieldsFunc(v, func(r rune) bool {
		return r == ',' || r == '，' || r == ';' || r == '；'
	}) {
		if part = strings.TrimSpace(part); part != "" {
			*a = append(*a, part)
		}
	}
	return nil
}

const helpText = `用法：
  mailer -c <配置文件> -f <文件> [选项]       # 读取本地文件：表格/Office文档/文本
  mailer -c <配置文件> -b <正文> -s <主题> [选项]  # AI/你已总结好的正文，直接发送
  mailer -x -f <文件>                       # 只把文件内容转成纯文本输出(供 AI 读取总结)
  mailer --help

正文来源（二选一或组合）：
  -f <文件>    读取本地文件。支持表格 .xlsx/.csv/.tsv，文档 .docx/.pptx，文本 .txt/.md/.log
               (Word/PowerPoint 会自动抽取文字；腾讯文档请先“导出”为 xlsx/docx 等)
  -b <正文>    直接传入已经总结好的正文（AI 总结模式推荐；可重复出现自动按换行拼接；
               传 -b - 表示从标准输入读取；传 -b @文件 表示从文件读取正文）

邮件信息：
  -t <邮箱>    收件人邮箱，可重复或 , ; ， ； 分隔传多个（如 -t "a@x.com,b@x.com"）。
               缺省用配置 defaults.to（默认 默认收件人（由 config defaults.to 或 -t 指定））；想先发给自己测试可直接
               -t 自己邮箱，无需改配置文件
  -s <主题>    邮件主题。省略时依次取：配置文件 defaults.subject(配置了才用) → 自动概括
               (正文/文件名首句，最多 30 字)；AI 模式建议显式传 -s 概括性主题
  -c <文件>    配置文件路径（JSON）。缺省查找环境变量 MAILER_CONFIG 与 ./config.json

其他选项：
  -x, --extract   只把文件内容转为纯文本打印到标准输出，不连 SMTP、不发送
                  (二进制文档 .docx/.pptx/.xlsx 也能转文本，供外部 AI Agent 读取后总结)
  --dry-run    只预览收件人/主题/正文等信息，不真正发送
  --verbose    打印详细日志（含 SMTP 会话与正文预览）
  -h, --help   显示本帮助（中文）并退出

退出码：
  0 成功（含 --dry-run / -x 成功）
  1 运行时错误（配置/文件/网络/认证等）
  2 命令行参数使用错误

默认业务规则（缺省收件人 SRE，可被 -t 或配置 defaults.to 覆盖）：
  收件人  默认收件人（由 config defaults.to 或 -t 指定），不抄送
  主题    默认概括文件内容；AI 模式由 AI 生成的一句话主题(-s)决定
  发信    使用配置文件里的企业邮箱账号（smtp.exmail.qq.com:465 SSL 为模板），
         凭据写配置文件或环境变量 MAILER_PASSWORD

AI 协作流程（配合 skills/daily-report-mailer.skill.md）：
  1) AI 用 “mailer -x -f <文件>” 取得文件纯文本(或自行读取)；
  2) AI 把内容总结成一段中文正文 + 生成一句话主题；
  3) AI 调用 “mailer -c config.json -t <收件人> -s <主题> -b <正文>”；
     收件人不传时默认发给 SRE。

示例：
  1) 预览本地文件自动汇总（不发送）：
     mailer -c config.json -f 新人磨合期记录.xlsx --dry-run
  2) 导出文件文本给 AI 阅读（不发送）：
     mailer -x -f 新人磨合期记录.xlsx
  3) 先发给自己测试（-t 覆盖收件人，无需改配置）：
     mailer -c config.json -t you@example.com -s "新人磨合期记录·汇总" -b "正文……" --dry-run
  4) 正式发给 SRE（缺省收件人）：
     mailer -c config.json -s "新人磨合期记录·汇总" -b "本周新人产出：共完成需求 12 个……"
  5) 中文帮助：
     mailer --help

配置文件模板见 config.example.json；完整说明见 README.md。
`

// ---------------------------------------------------------------------------
// 主流程
// ---------------------------------------------------------------------------

func printHelp(w io.Writer) {
	fmt.Fprintf(w, "%s v%s\n\n%s\n", appName, appVersion, helpText)
}

func printUsageError(w io.Writer, msg string) {
	fmt.Fprintf(w, "错误：%s\n\n执行 mailer --help 查看完整中文帮助。\n", msg)
}

func run(args []string, stdout, stderr io.Writer) int {
	fs := flag.NewFlagSet(appName, flag.ContinueOnError)
	fs.SetOutput(io.Discard) // 错误信息由我们自己以中文输出

	var opts cliOptions
	fs.StringVar(&opts.configPath, "c", "", "配置文件路径")
	fs.StringVar(&opts.filePath, "f", "", "源文件路径(表格/Office文档/文本)")
	fs.Var(&opts.bodyParts, "b", "直接传入正文（可重复，自动换行拼接）")
	fs.Var(&opts.toList, "t", "收件人邮箱(可重复或用,分隔；缺省=配置 defaults.to=SRE)")
	fs.Var(&opts.toList, "to", "收件人邮箱（同 -t）")
	fs.StringVar(&opts.subject, "s", "", "邮件主题（省略则自动概括文件内容）")
	fs.StringVar(&opts.subject, "subject", "", "邮件主题（同 -s）")
	fs.BoolVar(&opts.dryRun, "dry-run", false, "仅预览不发送")
	fs.BoolVar(&opts.extract, "x", false, "仅把文件内容转纯文本输出(给AI阅读)")
	fs.BoolVar(&opts.extract, "extract", false, "仅把文件内容转纯文本输出(同 -x)")
	fs.BoolVar(&opts.verbose, "verbose", false, "打印详细日志")
	fs.BoolVar(&opts.help, "help", false, "显示帮助")
	fs.BoolVar(&opts.help, "h", false, "显示帮助")

	if err := fs.Parse(args); err != nil {
		printUsageError(stderr, err.Error())
		return 2
	}
	if opts.help {
		printHelp(stdout)
		return 0
	}
	if len(fs.Args()) > 0 {
		printUsageError(stderr, fmt.Sprintf("存在无法识别的多余参数: %v（本工具不需要位置参数）", fs.Args()))
		return 2
	}

	// 0) -x/--extract：只导出文件纯文本给外部 AI 阅读（不需要配置文件/网络）
	if opts.extract {
		if opts.filePath == "" {
			printUsageError(stderr, "-x/--extract 需要配合 -f <文件> 使用")
			return 2
		}
		doc, err := readSourceFile(opts.filePath)
		if err != nil {
			fmt.Fprintf(stderr, "错误：%v\n", err)
			return 1
		}
		fmt.Fprint(stdout, docExtractText(doc))
		return 0
	}

	vlog := func(format string, a ...interface{}) {
		if opts.verbose {
			fmt.Fprintf(stderr, "[mailer] "+format+"\n", a...)
		}
	}

	// 1) 配置
	cfgPath := resolveConfigPath(opts.configPath)
	cfg, err := loadConfig(cfgPath, opts.dryRun)
	if err != nil {
		fmt.Fprintf(stderr, "错误：%v\n", err)
		return 1
	}
	if pwd := os.Getenv(envPassword); pwd != "" {
		cfg.Smtp.Password = pwd
		vlog("已从环境变量 %s 读取授权码", envPassword)
	}
	vlog("配置文件: %s（SMTP %s:%d，安全模式 %s）", cfgPath, cfg.Smtp.Host, cfg.Smtp.Port, cfg.Smtp.Security)

	// 2) 收件人/抄送：优先 -t/--to（可指定一个或多个，发给别人或先发自己测试），
	//    缺省用配置 defaults.to；未配置且未用 -t 时给出明确提示。不抄送
	recipients := cfg.Defaults.To
	if len(opts.toList) > 0 {
		recipients = []string(opts.toList)
	}
	hasRecipient := false
	for _, r := range recipients {
		if strings.TrimSpace(r) != "" {
			hasRecipient = true
			break
		}
	}
	if !hasRecipient {
		fmt.Fprintf(stderr, "错误：未指定收件人：请用 -t <邮箱> 指定，或在配置文件 defaults.to 中设置。\n")
		return 2
	}
	to, err := parseMailAddressList(recipients)
	if err != nil {
		fmt.Fprintf(stderr, "错误：收件人地址解析失败(-t/--to 或 defaults.to)：%v\n", err)
		return 1
	}
	cc, err := parseMailAddressList(cfg.Defaults.Cc)
	if err != nil {
		fmt.Fprintf(stderr, "错误：defaults.cc 解析失败：%v\n", err)
		return 1
	}

	// 3) 正文来源：-b 优先；其次 -f 自动摘要
	body, bodySrc, err := resolveBody(&opts)
	if err != nil {
		fmt.Fprintf(stderr, "错误：%v（请至少提供 -f <源文件> 或 -b <正文> 之一）\n\n执行 mailer --help 查看完整中文帮助。\n", err)
		return 1
	}

	// 3.5) 主题派生：-s 参数 > 配置 defaults.subject(配置了才用) > 自动概括(文件名/正文首句) > 兜底
	subject := strings.TrimSpace(opts.subject)
	if subject == "" {
		subject = strings.TrimSpace(cfg.Defaults.Subject)
	}
	if subject == "" {
		if len(opts.bodyParts) == 0 && opts.filePath != "" {
			subject = "《" + filepath.Base(opts.filePath) + "》内容汇总"
		} else if first := firstNonEmptyLine(body); first != "" {
			subject = truncate(first, 30)
		}
	}
	if subject == "" {
		subject = defaultSubject
	}
	vlog("主题: %s", subject)

	// 4) 预览输出
	var tos, ccs []string
	for _, a := range to {
		tos = append(tos, a.headerString())
	}
	for _, a := range cc {
		ccs = append(ccs, a.headerString())
	}
	fmt.Fprintf(stdout, "配置: %s\n", cfgPath)
	fmt.Fprintf(stdout, "收件人: %s\n", strings.Join(tos, ", "))
	if len(ccs) > 0 {
		fmt.Fprintf(stdout, "抄送: %s\n", strings.Join(ccs, ", "))
	}
	fmt.Fprintf(stdout, "主题: %s\n", subject)
	fmt.Fprintf(stdout, "正文来源: %s\n", bodySrc)
	fmt.Fprintf(stdout, "正文长度: %d 字\n", utf8.RuneCountInString(body))
	if opts.verbose {
		fmt.Fprintf(stdout, "---- 正文预览 ----\n%s\n---- 正文结束 ----\n", truncate(body, 400))
	}

	if opts.dryRun {
		fmt.Fprintln(stdout, "--dry-run：已通过全部本地校验，未连接 SMTP，未发送任何邮件。")
		return 0
	}

	// 5) 发送
	fmt.Fprintf(stdout, "正在通过 %s:%d(%s) 发送……\n", cfg.Smtp.Host, cfg.Smtp.Port, cfg.Smtp.Security)
	if err := sendMail(cfg, to, cc, subject, body, opts.verbose); err != nil {
		fmt.Fprintf(stderr, "错误：邮件发送失败：%v\n", err)
		return 1
	}
	fmt.Fprintf(stdout, "邮件发送成功：主题“%s”，收件人 %s。\n", subject, strings.Join(tos, ", "))
	return 0
}

// resolveBody 决定邮件正文：优先 -b，否则读 -f 文件自动摘要。
func resolveBody(opts *cliOptions) (body, sourceDesc string, err error) {
	if len(opts.bodyParts) > 0 {
		content := strings.Join(opts.bodyParts, "\n")
		// 单一 -b - ：从标准输入读取
		if len(opts.bodyParts) == 1 && content == "-" {
			data, rerr := io.ReadAll(io.LimitReader(os.Stdin, maxBodyBytes+1))
			if rerr != nil {
				return "", "", fmt.Errorf("读取标准输入失败: %w", rerr)
			}
			if int64(len(data)) > maxBodyBytes {
				return "", "", fmt.Errorf("标准输入内容超过 %d 字节上限", maxBodyBytes)
			}
			return strings.TrimRight(string(data), "\r\n"), "标准输入(-b -)", nil
		}
		// 单一 -b @文件：从文件读正文
		if len(opts.bodyParts) == 1 && strings.HasPrefix(content, "@") && len(content) > 1 {
			p := strings.TrimSpace(content[1:])
			data, rerr := readBodyFile(p)
			if rerr != nil {
				return "", "", rerr
			}
			return strings.TrimRight(string(data), "\r\n"), fmt.Sprintf("正文文件(-b @%s)", p), nil
		}
		if int64(len(content)) > maxBodyBytes {
			return "", "", fmt.Errorf("-b 正文超过 %d 字节上限", maxBodyBytes)
		}
		return content, "-b 参数(命令行正文)", nil
	}

	if opts.filePath != "" {
		doc, rerr := readSourceFile(opts.filePath)
		if rerr != nil {
			return "", "", rerr
		}
		bodyText := composeAutoBody(doc, filepath.Base(opts.filePath))
		return bodyText, fmt.Sprintf("自动摘要(-f %s)", opts.filePath), nil
	}

	return "", "", errors.New("缺少正文来源")
}

func readBodyFile(p string) ([]byte, error) {
	fi, err := os.Stat(p)
	if err != nil {
		return nil, fmt.Errorf("正文文件 %q 无法访问: %w", p, err)
	}
	if fi.IsDir() {
		return nil, fmt.Errorf("正文路径 %q 是目录", p)
	}
	if fi.Size() > maxBodyBytes {
		return nil, fmt.Errorf("正文文件 %q 超过 %d 字节上限", p, maxBodyBytes)
	}
	f, err := os.Open(p)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	data, err := io.ReadAll(io.LimitReader(f, maxBodyBytes+1))
	if err != nil {
		return nil, err
	}
	return data, nil
}

func main() {
	os.Exit(run(os.Args[1:], os.Stdout, os.Stderr))
}
