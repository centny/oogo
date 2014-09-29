package oogo

import (
	"fmt"
	"testing"
)

func TestOO(t *testing.T) {
	Init()
	defer Destory()
	calc, err := NewCalc()
	if err != nil {
		t.Error(err.Error())
		return
	}
	ss, err := calc.SheetI(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	for i := 0; i < 10; i++ {
		for j := 0; j < 5; j++ {
			ss.SetV(i, j, float64(i*j*6))
		}
	}
	ss.SetText(11, 0, "abc1")
	ss.SetText(11, 1, "这是中文,chinese")
	ss.SetFormula(11, 2, "=D4+E5")
	calc.Store(XLSX, "file:///C:/Users/Cny/Desktop/oo.xlsx")
	//
	fmt.Println("----------------->")
	//
	calc.Close()
	//
	calc, err = OpenCalc("file:///C:/Users/Cny/Desktop/oo.xlsx")
	ss, err = calc.SheetI(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(ss.GetText(11, 0))
	fmt.Println(ss.GetText(11, 1))
	fmt.Println(ss.GetText(11, 2))
	fmt.Println(ss.GetV(11, 2))
	fmt.Println(ss.GetFormula(11, 2))
	//
	ss, err = calc.NewSheet("Abc", 1)
	if err != nil {
		t.Error(err.Error())
		return
	}
	ss.SetText(0, 0, "abc1")
	ss.SetText(0, 1, "这是中文,chinese")
	ss, err = calc.SheetN("Abc")
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(ss.GetText(0, 0))
	fmt.Println(ss.GetText(0, 1))
	//
	calc.Store(XLSX, "file:///C:/Users/Cny/Desktop/oo2.xlsx")
	//
	calc.Close()
}
func TestFileProtocol(t *testing.T) {
	fmt.Println(FileProtocolPath("~"))
	fmt.Println(FileProtocolPath("sfdsf"))
	fmt.Println(FileProtocolPath("/sdfs/sfdsf"))
	fmt.Println(FileProtocolPath("C:\\s\\sdfs"))
}
