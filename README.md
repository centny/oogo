LibreOffice-GO
======
> library for calling LibreOffice by golang

===
#Install
#####＞Install LibreOfice and LibreOfice SDK,compile idl.

* Install mingw on window(only for window)
* Install LibreOffice and LibreOffice sdk from <http://www.libreoffice.org/download/>(install sdk to location withou blank space in the path on window,like C:\LibreOffice4)
* Compile idl to generate C++ header:

linux/unix/osx:

```
cd <LibreOffice4.3_SDK path>
./setsdkenv_windows
cppumaker -Gc -O./inc <LibreOffice Path>/misc/types.rdb <LibreOffice Path>/types/offapi.rdb

```

windows:

```
cd <LibreOffice4.3_SDK path>

setsdkenv_windows

cppumaker -Gc -O.\inc <LibreOffice Path>\URE\misc\types.rdb <LibreOffice Path>\program\types\offapi.rdb

```


#####＞GO get(not install):

```
go get -d github.com/Centny/oogo
```

#####＞Compile the liboogo

linux/unix/osx:

```
cd $GOPATH/src/github.com/Centny/oogo
./autoget.sh
./configure --prefix=/usr/local

//for osx
make "CPPFLAGS=-I<LibreOffice4.3_SDK path>/include -I<LibreOffice4.3_SDK path>/inc -DUNX -DGCC -DMACOSX -DCPPU_ENV=s5abi"

//for linux

make install
```

windows:

```
1.go to %GOPATH%/src/github.com/Centny/oogo/liboogo
2.open liboogo.vcxproj
3.update include and library directory configure to <LibreOffice4.3_SDK path>,default is C:\LibreOffice4\sdk
4.build the oogo.lib and oogo.dll

```

#####＞Go install

linux/unix/osx:

```
export CGO_LDFLAGS="-L<LibreOffice4.3_SDK path>/<platform>/lib -L/usr/local/lib -loogo -luno_cppu -luno_cppuhelpergcc3 -luno_purpenvhelpergcc3 -luno_sal -luno_salhelpergcc3"

go install github.com/Centny/oogo	
```

windows:

```
set CGO_LDFLAGS=-L<LibreOffice4.3_SDK path>\lib -L%GOPATH%\src\github.com\Centny\oogo -loogo -licppu -licppuhelper -lipurpenvhelper -lisal -lisalhelper

go install github.com/Centny/oogo
```


#Example

```go

	oogo.Init()                 //initial the environment.
	defer oogo.Destory()        //destory all.
	calc, err := oogo.NewCalc() //new Spreadsheet
	// calc, err := oogo.OpenCalc("file://tmp/oo.xlsx")
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
	calc.Store(oogo.XLSX, "file:///tmp/oo.xlsx") //save
```