package hf

import (
	"strings"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"fmt"
	"bytes"
	"strconv"
	"github.com/qjsoftcn/gutils"
	"os"
	"encoding/json"
	"log"
	"github.com/qjsoftcn/confs"
	"path/filepath"
)

const (
	//单元格最大显示字符数
	maxDisplayCharsNum = 80
	//pt to px mutiple num
	heightMutipleNum = 1.35
	//内外宽度倍数
	widthMutipleNum = 6.1
)

const (
	tabs_t_path  = "xlsx_t/tabs.t"
	tabs_td_path = "xlsx_t/tabs_td.t"
	tabs_name    = "tabs.html"

	frame_t_path = "xlsx_t/f.t"
	frame_name   = "index.html"
)

var mid_url=confs.GetString("rt","mid","url")
var bottom_url=confs.GetString("rt","bottom","url")
var head_url=confs.GetString("rt","bottom","url")

func SetUrl(){

}

//make rt tabs
func makeTabs(sheets []*xlsx.Sheet, destDir string) string {
	fc, err := ioutil.ReadFile(tabs_t_path)
	if err != nil {
		fmt.Println("read tab template ", err)
	}

	ftdc, err := ioutil.ReadFile(tabs_td_path)
	if err != nil {
		fmt.Println("read tabs td template ", err)
	}

	rs := bytes.Runes(fc)
	t := string(rs)

	rs_ftdc := bytes.Runes(ftdc)
	t_ftdc := string(rs_ftdc)

	sheetTabs := ""
	selectedFile := "s1"
	for index, sheet := range sheets {
		st := t_ftdc
		fn := "s" + strconv.Itoa(index+1)

		if sheet.Hidden {
			st = strings.Replace(st, "${dis}", "style='display:none;'", -1)
		} else {
			st = strings.Replace(st, "${dis}", gutils.EmptyString, -1)
		}

		st = strings.Replace(st, "${sheetName}", sheet.Name, -1)

		url := "\"" + mid_url + fn + "\""
		st = strings.Replace(st, "${sheetFile}", url, -1)

		sheetTabs += st + "\n"
		if sheet.Selected {
			selectedFile = fn
		}

	}
	//替换sheetTabs
	t = strings.Replace(t, "${sheetTabs}", sheetTabs, -1)

	tabsPath := filepath.Join(destDir, tabs_name)
	ioutil.WriteFile(tabsPath, []byte(t), 0777)
	return selectedFile
}

func makeFrame(selectedFile, destDir string) {
	fc, err := ioutil.ReadFile(frame_t_path)
	if err != nil {
		fmt.Println("read frame ", err)
	}

	rs := bytes.Runes(fc)
	t := string(rs)

	t = strings.Replace(t, "${head_url}", head_url, -1)
	t = strings.Replace(t, "${mid_url}", mid_url+selectedFile, -1)
	t = strings.Replace(t, "${bottom_url}", bottom_url, -1)

	f_html := filepath.Join(destDir, frame_name)
	ioutil.WriteFile(f_html, []byte(t), 0777)
}

func RtDir(rname string) string {
	dir := dst + "/" + rname
	if gutils.PathExists(dir) {
		return dir
	}
	//mk dir
	err := os.MkdirAll(dir, 0777)
	if err != nil {
		log.Println(err)
	}
	return dir
}

const (
	conf_group = "report"
	rt_dir     = "rtRoot"
	dot        = "."
	underline  = "_"
	sheetShort = "s"
)

var dst string = confs.GetString(conf_group, rt_dir)

//xlsx to html
//xlsxFile is xlsx file path
func XlsxToHtml(xlsxFile string, destDir, htmlFileName string) (bool, error) {
	//read excels
	xlFile, err := xlsx.OpenFile(xlsxFile)
	if err != nil {
		log.Fatal(err)
		return false, err
	}

	//make template is target
	rname := gutils.GetFileName(xlsxFile)
	//create destdir
	destDir = gutils.Dir(destDir)

	//make sheet tabs and frameSet html page
	sf := makeTabs(xlFile.Sheets, rname)
	makeFrame(sf, rname)

	var sheetIndex string

	calcCells := make([]CalcCell, 0)
	for si, sheet := range xlFile.Sheets {
		sheetIndex = fmt.Sprint(sheetShort, si+1)
		calcCells = makeSheet(rname, sheetIndex, sheet, calcCells)
	}

	dir := RtDir(rname)

	fn := dir + "/conf.json"
	jcs, err := json.Marshal(calcCells)
	if err != nil {
		log.Println("make ", rname, " json conf ", err)
	} else {
		ioutil.WriteFile(fn, []byte(jcs), 0777)
	}

}

const (
	rht_sheet_path = "rht/sheet.t"
)

type CalcCellType int

const (
	CCTypeConst   = iota
	CCTypeFormula
)

type CalcCell struct {
	Key  string
	Cell interface{}
	Type CalcCellType
}

func (this CalcCell) IsConst() bool {
	return this.Type == CCTypeConst
}

func (this CalcCell) IsFormula() bool {
	return this.Type == CCTypeFormula
}

func NewCalcCell(key string, cell interface{}, varType CalcCellType) CalcCell {
	cc := CalcCell{}
	cc.Type = varType
	cc.Key = key
	cc.Cell = cell
	return cc
}

func makeSheet(rname, sheetIndex string, sheet *xlsx.Sheet, calcCells []CalcCell) []CalcCell {

	var buffer bytes.Buffer
	fc, err := ioutil.ReadFile(rht_sheet_path)
	if err != nil {
		log.Println("read tab template ", err)
	}

	rs := bytes.Runes(fc)
	t := string(rs)

	tabW := 0.0
	for _, col := range sheet.Cols {
		if col.Width == 0 {
			tabW += xlsx.ColWidth
		} else {
			tabW += col.Width
		}

	}

	osw := fmt.Sprintf("%.0f", tabW*heightMutipleNum*widthMutipleNum)
	isw := fmt.Sprintf("%.0f", tabW*widthMutipleNum)

	t = strings.Replace(t, "${sheet.ow}", osw, 1)
	t = strings.Replace(t, "${sheet.iw}", isw, 1)

	dRH := sheet.SheetFormat.DefaultRowHeight
	maxCol := sheet.MaxCol

	sa := gutils.NewSheetAxiser()

	for ri, row := range sheet.Rows {
		buffer.WriteString("<tr ")
		if row.Hidden {
			buffer.WriteString("style='display:none'")
		} else {
			rh := row.Height
			if rh == 0 {
				rh = dRH
			}
			buffer.WriteString("height=" + fmt.Sprintf("%.0f", rh*heightMutipleNum))
			buffer.WriteString(" style='")
			buffer.WriteString("height:" + fmt.Sprintf("%.2f", rh) + "pt'")
		}

		buffer.WriteString(">\n")
		cellSize := len(row.Cells)

		if cellSize == 0 {
			//no cells
			for i := 0; i < maxCol; i++ {
				buffer.WriteString("<td></td>")
			}
		} else {
			for ci, cell := range row.Cells {
				rowSpan := cell.VMerge
				colSpan := cell.HMerge
				if rowSpan > 0 || colSpan > 0 {
					sa.Span(ri, ci, rowSpan, colSpan)
				}
				if !sa.IsSpan(ri, ci) {
					td, id, calc, isFormula, valid := cell2str(sheetIndex, sa.Axis(ri, ci, false), rowSpan, colSpan, *cell)
					if valid {
						if isFormula {
							cc := NewCalcCell(sheetIndex+underline+strings.ToLower(id), calc, CCTypeFormula)
							calcCells = append(calcCells, cc)
						} else {
							cc := NewCalcCell(sheetIndex+underline+strings.ToLower(id), calc, CCTypeConst)
							calcCells = append(calcCells, cc)
						}
					}
					buffer.WriteString(td)

				}
			}
		}

		buffer.WriteString("</tr>\n")
	}

	dir := RtDir(rname)
	fn := dir + "/" + sheetIndex
	t = strings.Replace(t, "${sheet.rows}", buffer.String(), 1)
	ioutil.WriteFile(fn+".html", []byte(t), 0777)

	log.Println(fn, "ok")
	return calcCells
}

func delSpan(ri, ci, rs, cs int, span map[string]bool) map[string]bool {

	for r := ri; r <= ri+rs; r++ {
		for c := ci; c <= ci+cs; c++ {

			if r == ri && c == ci {
				continue
			} else {
				//非起始位置
				key := strconv.Itoa(r) + "_" + strconv.Itoa(c)
				span[key] = true

			}

		}
	}
	return span
}

func isSpan(ri, ci int, span map[string]bool) bool {
	key := strconv.Itoa(ri) + "_" + strconv.Itoa(ci)
	if v, ok := span[key]; ok {

		return v
	}

	return false
}

func cell2str(sheetIndex, cellId string, vm, hm int, cell xlsx.Cell) (string, string, interface{}, bool, bool) {

	var buffer bytes.Buffer
	buffer.WriteString("<td ")

	buffer.WriteString(" id='" + sheetIndex + "_" + cellId + "'")
	if vm > 0 {
		buffer.WriteString(" rowSpan=" + strconv.Itoa(vm+1))
	}
	if hm > 0 {
		buffer.WriteString(" colSpan=" + strconv.Itoa(hm+1))
	}

	bs := []byte(cell.Value)
	rs := bytes.Runes(bs)
	vl := len(rs)

	cellView, calcCell, needTitle, isFormula, isVaild := delCellView(rs, cell)

	if cell.Hidden {
		buffer.WriteString(" style='display:none;'")
	} else {
		buffer.WriteString(cellStyleStr(*cell.GetStyle(), vl))
	}

	if needTitle {
		buffer.WriteString(" title=\"" + cell.Value + "\"")
	}

	buffer.WriteString(">")

	buffer.WriteString(cellView)
	buffer.WriteString("</td>\n")

	if isFormula {
		//need check funcs
		cc, _ := calcCell.(string)
		calcCell = funcs.Check(sheetIndex, cc)
	}

	return buffer.String(), cellId, calcCell, isFormula, isVaild

}

const (
	calc_cell_prefix = ':'
	calc_tip         = "..."
)

func delCellView(rs []rune, cell xlsx.Cell) (string, interface{}, bool, bool, bool) {
	vl := len(rs)
	if vl == 0 {
		return "", "", false, false, false
	}

	if rs[0] == calc_cell_prefix {
		//calc cell
		if vl > 1 {
			return calc_tip, string(rs[1:]), true, true, true
		} else {
			return string(rs), "", false, false, false
		}
	} else {
		var calcCell interface{}
		switch cell.Type() {
		case xlsx.CellTypeBool:
			calcCell = cell.Bool()
		case xlsx.CellTypeDate:
			calcCell, _ = cell.GetTime(true)
		case xlsx.CellTypeNumeric:
			calcCell, _ = cell.Float()
		default:
			calcCell = string(rs)
		}

		if vl > maxDisplayCharsNum {
			return string(rs[:maxDisplayCharsNum]) + calc_tip, calcCell, true, false, true
		} else {
			return string(rs), calcCell, false, false, true
		}
	}
}

func cellStyleStr(style xlsx.Style, vl int) string {
	s := ""

	if style.ApplyBorder {
		s += cellBorderStr("left", style.Border.Left, style.Border.LeftColor)
		s += cellBorderStr("right", style.Border.Right, style.Border.RightColor)
		s += cellBorderStr("top", style.Border.Top, style.Border.TopColor)
		s += cellBorderStr("bottom", style.Border.Bottom, style.Border.BottomColor)
	}

	if style.ApplyFont {
		s += cellFontStr(style.Font)
	}

	if style.ApplyFill {
		s += cellBgStr(style.Fill)
	}

	if style.ApplyAlignment {
		s += cellAlStr(style.Alignment)
	}

	if vl > maxDisplayCharsNum {
		s += "white-space:nowrap;overflow:hidden;word-break:keep-all;"
	}

	if len(s) > 0 {
		s = " style='" + s + "'"
		return s
	}

	return s
}

func cellAlStr(al xlsx.Alignment) string {
	s := ""
	if len(al.Horizontal) > 0 {
		s += "text-align:" + al.Horizontal + ";"
	}

	if len(al.Vertical) > 0 {
		s += "vertical-align:" + al.Vertical + ";"
	}

	if al.WrapText {
		s += "word-break:break-all;"
	}

	return s
}

func cellFontStr(font xlsx.Font) string {
	s := ""

	s += "font-size:" + strconv.Itoa(font.Size) + ";"
	s += "font-family:" + font.Name + ";"

	if len(font.Color) > 0 {
		s += "color:#" + font.Color[2:] + ";"
	}

	if font.Italic {
		s += "font-style:italic;"
	}

	if font.Bold {
		s += "font-weight:700;"
	}

	if font.Underline {
		s += "text-decoration:underline;text-underline-style:single;"
	}

	return s

}

func cellBorderStr(f, b, c string) string {
	s := ""

	if strings.Compare(strings.ToLower(b), "thin") == 0 {
		s += "border-" + f + ":.5pt solid;"
	}

	if len(c) > 0 {
		s += "border-" + f + "-color:#" + c[2:] + ";"
	}

	return s
}

func cellBgStr(fill xlsx.Fill) string {

	s := ""

	if len(fill.FgColor) > 0 {
		s += "background:#" + fill.FgColor[2:] + ";"
	}

	return s
}
