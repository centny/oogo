LibreOffice-GO
======
> library for calling LibreOffice/OpenOffice by golang

===
#Install

* Install LibreOffice and LibreOffice SDK.
* Install GCC (using mingw on window).
* GO get:

linux/unix/osx:

 ```
cd <LibreOffice4.3_SDK path>/

cppumaker -Gc -O./inc <LibreOffice Path>/misc/types.rdb <LibreOffice Path>/types/offapi.rdb

export CGO_CPPFLAGS="-I<LibreOffice4.3_SDK path>/include/ -I<LibreOffice4.3_SDK path>/inc"

export CGO_LDFLAGS="-L<LibreOffice4.3_SDK path>/<platform>/lib -luno_cppu -luno_cppuhelpergcc3 -luno_purpenvhelpergcc3 -luno_sal -luno_salhelpergcc3"

go get github.com/Centny/oogo	
```

win32:

```
cd <LibreOffice4.3_SDK path>/

cppumaker -Gc -O./inc <LibreOffice Path>/misc/types.rdb <LibreOffice Path>/types/offapi.rdb

set CGO_CFLAGS=-I<LibreOffice4.3_SDK path>\include -I%JAVA_HOME%\include\win32

set CGO_LDFLAGS=-L%JAVA_HOME%\lib -ljvm

go get github.com/Centny/jnigo	
```


#Example

```go

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
	calc.Store(XLSX, "file:///tmp/oo.xlsx") //save
```