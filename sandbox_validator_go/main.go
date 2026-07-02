package main

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type FieldDef struct {
	Col   int
	Name  string
	Type  string
	Req   bool
	Extra interface{}
}

type Err struct {
	Sheet string
	Row   int
	Col   int
	Name  string
	Value string
	Desc  string
}

func (e Err) String() string {
	v := e.Value
	if len(v) > 25 {
		v = v[:25]
	}
	if v == "" {
		v = "(空)"
	}
	return fmt.Sprintf("行%d [%s] '%s' → %s", e.Row, e.Name, v, e.Desc)
}

var storeFields = []FieldDef{
	{0, "省份", "enum", true, []string{"海南省", "广东省", "广西壮族自治区", "云南省", "四川省", "河北省", "山西省", "辽宁省", "吉林省", "黑龙江省", "江苏省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "河南省", "湖北省", "湖南省", "海南省", "重庆市", "新疆维吾尔自治区", "宁夏回族自治区", "内蒙古自治区", "广西壮族自治区", "陕西省", "甘肃省", "青海省", "贵州省", "吉林省", "黑龙江省", "北京市", "上海市", "天津市", "台湾省", "香港特别行政区", "澳门特别行政区"}},
	{1, "城市", "text", true, nil},
	{2, "区县", "text", true, nil},
	{3, "乡镇", "text", true, nil},
	{4, "进度", "enum", true, []string{"已存在", "无规划", "洽谈中"}},
	{5, "编码", "code", true, nil},
	{6, "名称", "text", true, nil},
	{7, "地理编码", "ref", true, "geo"},
	{8, "商圈编码", "ref", true, "biz"},
	{9, "类型", "enum", true, []string{"综合卖场", "荣耀体验店", "荣耀授权专卖店", "耀生活", "华为授权体验店", "华为堡垒店", "华为非授权店", "OPPO体验店", "VIVO体验店", "小米体验店", "苹果体验店", "其他品牌专卖店", "运营商厅店", "NKA综合卖场", "其他"}},
	{10, "capa", "uint", true, nil},
	{11, "面积", "uint", false, nil},
	{12, "状态", "enum", true, []string{"不合作", "营业中", "闭店", "装修中"}},
	{13, "运营商", "enum", true, []string{"公开", "公开&三网", "排它-移动", "排它-电信", "排它-联通", "非排它-移动电信", "非排它-移动联通", "非排它-电信联通"}},
	{14, "联系人", "text", true, nil},
	{15, "联系电话", "phone", true, nil},
	{16, "地址", "text", true, nil},
	{17, "渠道编码", "chan", true, nil},
	{18, "客户名称", "text", true, nil},
	{19, "0-2500", "uint", true, nil},
	{20, "2500-4000", "uint", true, nil},
	{21, ">4000", "uint", true, nil},
	{22, "荣耀", "uint", true, nil},
	{23, "华为", "uint", true, nil},
	{24, "小米", "uint", true, nil},
	{25, "VIVO", "uint", true, nil},
	{26, "苹果", "uint", true, nil},
	{27, "OPPO", "uint", true, nil},
	{28, "荣耀PC", "uint", true, nil},
	{29, "华为PC", "uint", true, nil},
	{30, "荣耀平板", "uint", true, nil},
	{31, "华为平板", "uint", true, nil},
	{32, "PC", "uint", true, nil},
	{33, "平板", "uint", true, nil},
	{34, "手表", "uint", true, nil},
	{35, "荣耀阵地数", "uint", false, nil},
	{36, "荣耀阵地位", "pos", false, nil},
	{37, "荣耀门头数", "uint", false, nil},
	{38, "华为阵地数", "uint", false, nil},
	{39, "华为阵地位", "pos", false, nil},
	{40, "华为门头数", "uint", false, nil},
	{41, "小米阵地数", "uint", false, nil},
	{42, "小米阵地位", "pos", false, nil},
	{43, "小米门头数", "uint", false, nil},
	{44, "VIVO阵地数", "uint", false, nil},
	{45, "VIVO阵地位", "pos", false, nil},
	{46, "VIVO门头数", "uint", false, nil},
	{47, "苹果营地数", "uint", false, nil},
	{48, "苹果阵地位", "pos", false, nil},
	{49, "苹果门头数", "uint", false, nil},
	{50, "OPPO阵地数", "uint", false, nil},
	{51, "OPPO阵地位", "pos", false, nil},
	{52, "OPPO门头数", "uint", false, nil},
	{53, "荣耀促销员", "uint", true, nil},
	{54, "华为促销员", "uint", true, nil},
	{55, "小米促销员", "uint", true, nil},
	{56, "VIVO促销员", "uint", true, nil},
	{57, "苹果促销员", "uint", true, nil},
	{58, "OPPO促销员", "uint", true, nil},
	{59, "客户员工", "uint", true, nil},
	{60, "店端人数", "uint", true, nil},
	{61, "国补", "enum", true, []string{"是", "否"}},
	{62, "城乡类别", "enum", true, []string{"主城区", "非主城区"}},
	{63, "街边/MALL", "enum", true, []string{"街边店", "MALL店"}},
	{64, "备注", "text", false, nil},
}

var customerFields = []FieldDef{
	{4, "跨省份", "enum", true, []string{"是", "否"}},
	{5, "跨地市", "enum", true, []string{"是", "否"}},
	{8, "圈层", "text", true, nil},
	{9, "华为属性", "enum", true, []string{"金种子", "TOP368", "TOP1000", "TOP1500", "NDKA+", "TOP2000", "TOP5000", "其他"}},
	{10, "OPPO层级", "enum", true, []string{"T+", "T", "S", "A", "其他"}},
	{11, "vivo层级", "enum", true, []string{"黑金+", "黑金", "钻石", "铂金", "其他"}},
	{12, "小米层级", "enum", true, []string{"蓝血Ultra", "蓝血Plus", "蓝血", "金牌", "卓越Ultra", "卓越Plus", "卓越", "优秀", "其他"}},
	{13, "运营商属性", "enum", true, []string{"公开", "公开&三网", "排它-移动", "排它-电信", "排它-联通", "非排它-移动电信", "非排它-移动联通", "非排它-电信联通"}},
	{14, "友商代理", "enum", true, []string{"是", "否"}},
	{15, "车授权", "enum", true, []string{"是", "否"}},
	{16, "老板名", "text", true, nil},
	{17, "老板电话", "phone", true, nil},
	{18, "操盘手名", "text", true, nil},
	{19, "操盘手电话", "phone", true, nil},
	{42, "0-2500容量", "uint", true, nil},
}

// RefData holds reference code sets
type RefData struct {
	GeoCodes map[string]bool
	BizCodes map[string]bool
}

func toInt(v interface{}) *int {
	if v == nil {
		return nil
	}
	switch val := v.(type) {
	case int:
		return &val
	case int64:
		i := int(val)
		return &i
	case float64:
		i := int(val)
		return &i
	case string:
		s := strings.TrimSpace(val)
		if s == "" || s == "nan" || s == "N/A" {
			return nil
		}
		s = strings.ReplaceAll(s, ",", "")
		i, err := strconv.Atoi(s)
		if err != nil {
			f, err2 := strconv.ParseFloat(s, 64)
			if err2 != nil {
				return nil
			}
			i = int(f)
		}
		return &i
	}
	return nil
}

func isBlank(v interface{}) bool {
	if v == nil {
		return true
	}
	s, ok := v.(string)
	if !ok {
		return false
	}
	return strings.TrimSpace(s) == ""
}

func safeStr(v interface{}) string {
	if v == nil {
		return ""
	}
	s := strings.TrimSpace(fmt.Sprintf("%v", v))
	if s == "nan" || s == "N/A" {
		return ""
	}
	return s
}

func cleanPhone(v string) string {
	re := regexp.MustCompile(`[^0-9]`)
	return re.ReplaceAllString(v, "")
}

func validateEnum(name string, val interface{}, eset []string) *Err {
	if isBlank(val) {
		return &Err{Name: name, Value: "", Desc: "该字段必填"}
	}
	sv := safeStr(val)
	for _, e := range eset {
		if sv == e {
			return nil
		}
	}
	return &Err{Name: name, Value: sv, Desc: fmt.Sprintf("无效值: %s", strings.Join(eset, "/"))}
}

func validateUint(name string, val interface{}, req bool) *Err {
	if isBlank(val) {
		if req {
			return &Err{Name: name, Value: "", Desc: "该字段必填"}
		}
		return nil
	}
	iv := toInt(val)
	if iv == nil {
		return &Err{Name: name, Value: safeStr(val), Desc: "非有效整数"}
	}
	if *iv < 0 || *iv > 999999 {
		return &Err{Name: name, Value: safeStr(val), Desc: fmt.Sprintf("超出范围: %d", *iv)}
	}
	return nil
}

func validatePhone(name string, val interface{}, req bool) *Err {
	if isBlank(val) {
		if req {
			return &Err{Name: name, Value: "", Desc: "该字段必填"}
		}
		return nil
	}
	sv := safeStr(val)
	if sv == "无" {
		return nil
	}
	s := cleanPhone(sv)
	if !regexp.MustCompile(`^1[3-9]\d{9}$`).MatchString(s) {
		return &Err{Name: name, Value: sv, Desc: "无效手机号"}
	}
	return nil
}

func validateCode(name string, val interface{}, plan string, req bool) *Err {
	if isBlank(val) {
		if req && plan == "已存在" {
			return &Err{Name: name, Value: "", Desc: "已存在门店必须填写编码"}
		}
		return nil
	}
	s := safeStr(val)
	if s == "无规划" || s == "洽谈中" || s == "/" || s == "无" {
		return nil
	}
	if regexp.MustCompile(`^[A-Za-z]`).MatchString(s) ||
		regexp.MustCompile(`^46\d{6,}$`).MatchString(s) {
		return nil
	}
	return &Err{Name: name, Value: s, Desc: "编码格式异常"}
}

func validatePos(name string, val interface{}) *Err {
	if isBlank(val) {
		return nil
	}
	sv := safeStr(val)
	if sv == "0" || sv == "0.0" || sv == "" {
		return nil
	}
	valid := []string{"TOP1", "TOP2", "TOP3", "其他", "无阵地"}
	for _, v := range valid {
		if sv == v {
			return nil
		}
	}
	return &Err{Name: name, Value: sv, Desc: "无效枚举: TOP1/TOP2/TOP3/其他/无阵地"}
}

func validateRef(name string, val interface{}, refSet map[string]bool, req bool, label string) *Err {
	if isBlank(val) {
		if req {
			return &Err{Name: name, Value: "", Desc: "该字段必填"}
		}
		return nil
	}
	sv := safeStr(val)
	if !refSet[sv] {
		return &Err{Name: name, Value: sv, Desc: fmt.Sprintf("%s编码不存在于系统", label)}
	}
	return nil
}

func validateChan(name string, val interface{}, req bool) *Err {
	if isBlank(val) {
		if req {
			return &Err{Name: name, Value: "", Desc: "该字段必填"}
		}
		return nil
	}
	s := safeStr(val)
	if s == "无" || regexp.MustCompile(`^\d+$`).MatchString(s) {
		return nil
	}
	return &Err{Name: name, Value: s, Desc: "渠道编码应为纯数字或'无'"}
}

func validateStoreRow(row []string, ref *RefData, rowNum int) []Err {
	var errs []Err

	for _, f := range storeFields {
		var val string
		if f.Col < len(row) {
			val = row[f.Col]
		}
		var e *Err
		switch f.Type {
		case "enum":
			eset := f.Extra.([]string)
			e = validateEnum(f.Name, val, eset)
		case "uint":
			e = validateUint(f.Name, val, f.Req)
		case "phone":
			e = validatePhone(f.Name, val, f.Req)
		case "code":
			plan := ""
			if 4 < len(row) {
				plan = row[4]
			}
			e = validateCode(f.Name, val, plan, f.Req)
		case "ref":
			label := f.Extra.(string)
			var refSet map[string]bool
			if label == "geo" {
				refSet = ref.GeoCodes
			} else {
				refSet = ref.BizCodes
			}
			e = validateRef(f.Name, val, refSet, f.Req, label)
		case "chan":
			e = validateChan(f.Name, val, f.Req)
		case "pos":
			e = validatePos(f.Name, val)
		case "text":
			if f.Req && isBlank(val) {
				e = &Err{Name: f.Name, Value: "", Desc: "该字段必填"}
			}
		}
		if e != nil {
			e.Row = rowNum
			e.Col = f.Col
			e.Sheet = "通讯通路门店"
			errs = append(errs, *e)
		}
	}

	// 勾稽：门店capa = 价位段 = 品牌
	pSum := 0
	bSum := 0
	kVal := 0
	hasK := false
	for _, idx := range []int{19, 20, 21} {
		if idx < len(row) {
			v := toInt(row[idx])
			if v != nil {
				pSum += *v
			}
		}
	}
	for _, idx := range []int{22, 23, 24, 25, 26, 27} {
		if idx < len(row) {
			v := toInt(row[idx])
			if v != nil {
				bSum += *v
			}
		}
	}
	if 10 < len(row) {
		v := toInt(row[10])
		if v != nil {
			kVal = *v
			hasK = true
		}
	}
	isClosed := 12 < len(row) && row[12] == "闭店"
	if !(isClosed && pSum == 0 && bSum == 0) {
		if pSum != bSum {
			errs = append(errs, Err{
				Sheet: "通讯通路门店", Row: rowNum, Col: 19,
				Name: "价位段≠品牌", Value: fmt.Sprintf("%d≠%d", pSum, bSum),
				Desc: fmt.Sprintf("价位段(%d)≠品牌合计(%d)", pSum, bSum),
			})
		}
		if hasK && kVal != pSum {
			errs = append(errs, Err{
				Sheet: "通讯通路门店", Row: rowNum, Col: 10,
				Name: "capa≠价位段", Value: fmt.Sprintf("%d≠%d", kVal, pSum),
				Desc: fmt.Sprintf("门店capa(%d)≠价位段之和(%d)≠品牌合计(%d)", kVal, pSum, bSum),
			})
		}
	}

	// TOP位跨品牌可重复，无需唯一性检查
	return errs
}

func validateCustomerRow(row []string, rowNum int) []Err {
	var errs []Err

	for _, f := range customerFields {
		var val string
		if f.Col < len(row) {
			val = row[f.Col]
		}
		var e *Err
		switch f.Type {
		case "enum":
			eset := f.Extra.([]string)
			e = validateEnum(f.Name, val, eset)
		case "uint":
			e = validateUint(f.Name, val, f.Req)
		case "phone":
			e = validatePhone(f.Name, val, f.Req)
		case "text":
			if f.Req && isBlank(val) {
				e = &Err{Name: f.Name, Value: "", Desc: "该字段必填"}
			}
		}
		if e != nil {
			e.Row = rowNum
			e.Col = f.Col
			e.Sheet = "通讯客户沙盘"
			errs = append(errs, *e)
		}
	}
	return errs
}

func readRefData(f *excelize.File) *RefData {
	ref := &RefData{
		GeoCodes: make(map[string]bool),
		BizCodes: make(map[string]bool),
	}
	for _, sn := range f.GetSheetList() {
		rows, _ := f.GetRows(sn)
		if len(rows) == 0 {
			continue
		}
		isGeo := false
		isBiz := false
		for _, cell := range rows[0] {
			if strings.Contains(cell, "地理信息编码") || strings.Contains(cell, "地理编码") {
				isGeo = true
			}
			if strings.Contains(cell, "商圈编码") {
				isBiz = true
			}
		}
		if !isGeo && !isBiz {
			continue
		}
		for _, row := range rows[1:] {
			for _, cell := range row {
				if cell != "" {
					if isGeo {
						ref.GeoCodes[cell] = true
					}
					if isBiz {
						ref.BizCodes[cell] = true
					}
				}
			}
		}
	}
	return ref
}

func colLetter(col int) string {
	s := ""
	for col > 0 {
		col--
		s = string(rune('A'+col%26)) + s
		col /= 26
	}
	return s
}

func processFile(fp string) ([]Err, error) {
	fmt.Printf("\n═══ %s ═══\n", filepath.Base(fp))
	startTime := time.Now()

	f, err := excelize.OpenFile(fp)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %v", err)
	}
	defer f.Close()

	ref := readRefData(f)
	fmt.Printf("引用: %d地理, %d商圈\n", len(ref.GeoCodes), len(ref.BizCodes))

	var storeErrs []Err
	var custErrs []Err
	var storeRowCount int
	var custRowCount int

	// ── 校验通讯通路门店 ──
	storeSheetName := "通讯通路门店"
	if sheetIdx, _ := f.GetSheetIndex(storeSheetName); sheetIdx >= 0 {
		rows, _ := f.GetRows(storeSheetName)
		// 获取实际数据最大列（跳过空列扫描）
		storeMaxCol := 0
		for _, row := range rows {
			if len(row) > storeMaxCol {
				storeMaxCol = len(row)
			}
		}
		storeRowCount = len(rows) - 3 // 减去表头3行
		if storeRowCount < 0 {
			storeRowCount = 0
		}
		for i := 3; i < len(rows); i++ {
			rowErrs := validateStoreRow(rows[i], ref, i+1)
			storeErrs = append(storeErrs, rowErrs...)
		}
		// 写出校验结果到最后一列+1
		writeCheckColumn(f, storeSheetName, storeErrs, storeMaxCol, startTime)
	}

	// ── 校验通讯客户沙盘 ──
	custSheetName := "通讯客户沙盘"
	if sheetIdx, _ := f.GetSheetIndex(custSheetName); sheetIdx >= 0 {
		rows, _ := f.GetRows(custSheetName)
		custMaxCol := 0
		for _, row := range rows {
			if len(row) > custMaxCol {
				custMaxCol = len(row)
			}
		}
		custRowCount = len(rows) - 3
		if custRowCount < 0 {
			custRowCount = 0
		}
		for i := 3; i < len(rows); i++ {
			rowErrs := validateCustomerRow(rows[i], i+1)
			custErrs = append(custErrs, rowErrs...)
		}
		writeCheckColumn(f, custSheetName, custErrs, custMaxCol, startTime)
	}

	// ── 保存为 _校验结果.xlsx ──
	outPath := strings.TrimSuffix(fp, filepath.Ext(fp)) + "_校验结果.xlsx"
	err = f.SaveAs(outPath)
	if err != nil {
		return nil, fmt.Errorf("保存失败: %v", err)
	}

	elapsed := time.Since(startTime).Seconds()
	fmt.Printf("✅ 已保存: %s\n", filepath.Base(outPath))
	fmt.Printf("  用时: %.1fs | 通路: %d 条错误 | 客户: %d 条错误\n", elapsed, len(storeErrs), len(custErrs))

	allErrs := append(storeErrs, custErrs...)
	return allErrs, nil
}

// writeCheckColumn 将校验结果写入数据最后一列之后的新列
func writeCheckColumn(f *excelize.File, sheetName string, errs []Err, dataMaxCol int, startTime time.Time) {
	// 找出数据最后一列的列号（0-based），校验列在其后
	checkCol := dataMaxCol  // 0-based
	checkCol1 := checkCol + 1 // 1-based for Excel

	// 获取总行数
	rows, _ := f.GetRows(sheetName)
	totalRows := len(rows)

	// 创建黄色表头样式
	yellowStyleID, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "000000", Size: 11},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1},
		Border:    []excelize.Border{{Type: "thin", Color: "000000", Style: 1}},
	})
	// 绿色（正确）
	greenStyleID, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"C6EFCE"}, Pattern: 1},
	})
	// 红色（错误）
	redStyleID, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FFC7CE"}, Pattern: 1},
	})

	// 收集行级错误映射
	rowErrs := make(map[int][]Err)
	for _, e := range errs {
		rowErrs[e.Row] = append(rowErrs[e.Row], e)
	}

	// 先清理所有行的校验列（从第4行开始）
	for rowIdx := 3; rowIdx < totalRows; rowIdx++ {
		cell := cellRef(checkCol1, rowIdx+1)
		f.SetCellValue(sheetName, cell, "")
		f.SetCellStyle(sheetName, cell, cell, greenStyleID)
	}

	// 写入表头（第2行）
	headerCell := cellRef(checkCol1, 2)
	f.SetCellValue(sheetName, headerCell, "校验结果")
	f.SetCellStyle(sheetName, headerCell, headerCell, yellowStyleID)

	// 写入错误行
	for rowNum, rowErrList := range rowErrs {
		var texts []string
		for _, e := range rowErrList {
			texts = append(texts, e.String())
		}
		text := strings.Join(texts, " | ")
		if len(text) > 490 {
			text = text[:490]
		}
		text = "❌ " + text
		cell := cellRef(checkCol1, rowNum)
		f.SetCellValue(sheetName, cell, text)
		f.SetCellStyle(sheetName, cell, cell, redStyleID)
	}

	// 写入正确行（无错误的行=绿色）
	correctCount := 0
	for rowIdx := 3; rowIdx < totalRows; rowIdx++ {
		rowNum := rowIdx + 1
		if _, hasErr := rowErrs[rowNum]; !hasErr {
			// 第1列值作为标识
			identifier := ""
			if rowIdx < len(rows) && len(rows[rowIdx]) > 0 {
				identifier = fmt.Sprintf("%v", rows[rowIdx][0])
			}
			if identifier != "" {
				cell := cellRef(checkCol1, rowNum)
				f.SetCellValue(sheetName, cell, "✅ 正确")
				f.SetCellStyle(sheetName, cell, cell, greenStyleID)
				correctCount++
			}
		}
	}

	// 设置列宽
	f.SetColWidth(sheetName, colLetter(int(checkCol1)), colLetter(int(checkCol1)), 80)
}

// cellRef 生成 Excel 列字母+行号
func cellRef(col1 int, row int) string {
	return fmt.Sprintf("%s%d", colLetter(col1), row)
}

func scanDir(path string) []string {
	var files []string
	filepath.Walk(path, func(p string, info os.FileInfo, err error) error {
		if err != nil {
			return nil
		}
		if !info.IsDir() && strings.HasSuffix(strings.ToLower(p), ".xlsx") &&
			!strings.HasPrefix(info.Name(), "~") {
			files = append(files, p)
		}
		return nil
	})
	return files
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("通讯沙盘校验工具 v1.0 (Go)")
		fmt.Println("用法:")
		fmt.Println("  sandbox_validator_go <文件1.xlsx> [文件2.xlsx ...]")
		fmt.Println("  sandbox_validator_go -dir <文件夹路径>")
		os.Exit(0)
	}

	var files []string
	if os.Args[1] == "-dir" && len(os.Args) >= 3 {
		files = scanDir(os.Args[2])
		fmt.Printf("扫描: %d 个文件\n", len(files))
	} else {
		files = os.Args[1:]
	}

	totalStart := time.Now()
	totalErrs := 0

	for _, fp := range files {
		fp = strings.TrimSpace(fp)
		if fp == "" {
			continue
		}
		errs, err := processFile(fp)
		if err != nil {
			fmt.Printf("❌ %s: %v\n", filepath.Base(fp), err)
		} else {
			totalErrs += len(errs)
		}
	}

	fmt.Printf("\n%s\n", strings.Repeat("=", 50))
	fmt.Printf("完成: %d 文件, 共 %d 条错误 (%.1fs)\n",
		len(files), totalErrs, time.Since(totalStart).Seconds())
}
