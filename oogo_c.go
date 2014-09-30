package oogo

/*
#include <stdio.h>
#include <stdlib.h>
#include "oogo_c.h"
#cgo darwin CPPFLAGS: -I/Users/cny/LibreOffice4.3_SDK/include -I/Users/cny/LibreOffice4.3_SDK/inc -DUNX -DGCC -DMACOSX -DCPPU_ENV=s5abi -arch x86_64
#cgo darwin LDFLAGS: -L/Users/cny/LibreOffice4.3_SDK/macosx/lib -L/usr/local/lib -loogo -luno_cppu -luno_cppuhelpergcc3 -luno_purpenvhelpergcc3 -luno_sal -luno_salhelpergcc3
#cgo win LDFLAGS: -LC:\LibreOffice4\sdk\lib -LC:\GOPATH\src\github.com\Centny\oogo -loogo -licppu -licppuhelper -lipurpenvhelper -lisal -lisalhelper
*/
import "C"
import (
	"errors"
	"path/filepath"
	"strings"
	"unsafe"
)

const (
	E_BUF_LEN = 2048
)
const (
	XLS  = "MS Excel 97"
	XLSX = "Calc MS Excel 2007 XML"
)

//Call object to map LibreOffice XSpreadsheetDocument
type Calc C.calc_c

//Call object to map LibreOffice XSpreadsheet
type Sheet C.sheet_c

//initial the LibreOffice(bootstrap)
func Init() int {
	return int(C.oogo_init())
}

//free all.
func Destory() int {
	return int(C.oogo_destory())
}

//check the error message.
func Error() error {
	buf := make([]byte, E_BUF_LEN)
	C.oogo_cpy_ebuf((*C.char)(unsafe.Pointer(&buf[0])))
	bs := string(buf)
	if len(bs) < 1 {
		return nil
	} else {
		return errors.New(bs)
	}
}

//create one calc document.
func NewCalc() (Calc, error) {
	calc := C.oogo_new_calc()
	if calc.code == 0 {
		return Calc(calc), nil
	} else {
		return Calc{}, Error()
	}
}

//open one cacl document.
func OpenCalc(url string) (Calc, error) {
	curl := C.CString(url)
	defer C.free(unsafe.Pointer(curl))
	calc := C.oogo_open_calc(curl)
	if calc.code == 0 {
		return Calc(calc), nil
	} else {
		return Calc{}, Error()
	}
}

//close the calc document.
func (c Calc) Close() error {
	code := C.oogo_close_calc(C.calc_c(c))
	if code == 0 {
		return nil
	} else {
		return Error()
	}
}

//save document to file by file and file url.
func (c Calc) Store(filter, url string) error {
	cfilter := C.CString(filter)
	curl := C.CString(url)
	defer func() {
		C.free(unsafe.Pointer(cfilter))
		C.free(unsafe.Pointer(curl))
	}()
	code := C.oogo_store_calc(C.calc_c(c), cfilter, curl)
	if code == 0 {
		return nil
	} else {
		return Error()
	}
}

//new sheet by name and index.
func (c Calc) NewSheet(name string, idx int) (Sheet, error) {
	cname := C.CString(name)
	defer C.free(unsafe.Pointer(cname))
	sheet := C.oogo_sheet_new(C.calc_c(c), cname, C.int(idx))
	if sheet.code == 0 {
		return Sheet(sheet), nil
	} else {
		return Sheet{}, Error()
	}
}

//find sheet by index.
func (c Calc) SheetI(idx int) (Sheet, error) {
	sheet := C.oogo_sheet_i(C.calc_c(c), C.int(idx))
	if sheet.code == 0 {
		return Sheet(sheet), nil
	} else {
		return Sheet{}, Error()
	}
}

//find sheet by name.
func (c Calc) SheetN(name string) (Sheet, error) {
	cname := C.CString(name)
	defer C.free(unsafe.Pointer(cname))
	sheet := C.oogo_sheet_n(C.calc_c(c), cname)
	if sheet.code == 0 {
		return Sheet(sheet), nil
	} else {
		return Sheet{}, Error()
	}
}

// //
//set number value by index.
func (s Sheet) SetV(x, y int, num float64) error {
	code := C.oogo_sheet_set_v(C.sheet_c(s), C.int(x), C.int(y), C.double(num))
	if code == 0 {
		return nil
	} else {
		return Error()
	}
}

//get number value by index.
func (s Sheet) GetV(x, y int) (float64, error) {
	var num C.double = 0
	code := C.oogo_sheet_get_v(C.sheet_c(s), C.int(x), C.int(y), &num)
	if code == 0 {
		return float64(num), nil
	} else {
		return 0, Error()
	}
}

// //
//set the formula by index.
func (s Sheet) SetFormula(x, y int, val string) error {
	cval := C.CString(val)
	defer C.free(unsafe.Pointer(cval))
	code := C.oogo_sheet_set_formula(C.sheet_c(s), C.int(x), C.int(y), cval)
	if code == 0 {
		return nil
	} else {
		return Error()
	}
}

//get the formula by index.
func (s Sheet) GetFormula(x, y int) (string, error) {
	var l C.int = 0
	code := C.oogo_sheet_get_formula_l(C.sheet_c(s), C.int(x), C.int(y), &l)
	if code != 0 {
		return "", Error()
	}
	if l < 1 {
		return "", nil
	}
	buf := make([]byte, int(l))
	code = C.oogo_sheet_cpy_formula(C.sheet_c(s), C.int(x), C.int(y), (*C.char)(unsafe.Pointer(&buf[0])), C.int(l))
	if code == 0 {
		return string(buf), nil
	} else {
		return "", Error()
	}
}

// //
//set the text by index.
func (s Sheet) SetText(x, y int, val string) error {
	cval := C.CString(val)
	defer C.free(unsafe.Pointer(cval))
	code := C.oogo_sheet_set_text(C.sheet_c(s), C.int(x), C.int(y), cval)
	if code == 0 {
		return nil
	} else {
		return Error()
	}
}

//get the text by index.
func (s Sheet) GetText(x, y int) (string, error) {
	var l C.int = 0
	code := C.oogo_sheet_get_text_l(C.sheet_c(s), C.int(x), C.int(y), &l)
	if code != 0 {
		return "", Error()
	}
	if l < 1 {
		return "", nil
	}
	buf := make([]byte, int(l))
	code = C.oogo_sheet_cpy_text(C.sheet_c(s), C.int(x), C.int(y), (*C.char)(unsafe.Pointer(&buf[0])), C.int(l))
	if code == 0 {
		return string(buf), nil
	} else {
		return "", Error()
	}
}

//return the end cell index of column and row
func (s Sheet) EndRL() (int, int, error) {
	var c C.int = 0
	var r C.int = 0
	code := C.oogo_sheet_end_r_l(C.sheet_c(s), &c, &r)
	if code == 0 {
		return int(c), int(r), nil
	} else {
		return 0, 0, Error()
	}
}

//cover the local file path to file:// path.
func FileProtocolPath(t string) string {
	t = strings.Trim(t, " \t")
	if strings.HasPrefix(t, "file://") {
		return t
	}
	t, _ = filepath.Abs(t)
	t = strings.Replace(t, "\\", "/", -1)
	if strings.HasPrefix(t, "/") {
		return "file://" + t
	} else {
		return "file:///" + t
	}
}
