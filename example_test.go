package oogo

import (
	"fmt"
)

func ExampleOOgo() {
	Init()                 //initial the environment.
	defer Destory()        //destory all.
	calc, err := NewCalc() //new Spreadsheet
	// calc, err := OpenCalc("file://tmp/oo.xlsx")
	if err != nil {
		return
	}
	defer calc.Close()
	ss, err := calc.SheetI(0) //get sheet.
	if err != nil {
		return
	}
	for i := 0; i < 10; i++ {
		for j := 0; j < 5; j++ {
			ss.SetV(i, j, float64(i*j*6))
		}
	}
	ss.SetText(11, 0, "abc1")
	ss.SetFormula(11, 2, "=D4+E5")
	fmt.Println(ss.GetText(11, 0))
	fmt.Println(ss.GetFormula(11, 2))
	fmt.Println(ss.GetV(1, 1))
	calc.Store(XLSX, FileProtocolPath("file:///tmp/oo.xlsx")) //save
}
