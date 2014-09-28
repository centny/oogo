#if 0
#include <cppuhelper/bootstrap.hxx>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/lang/XMultiComponentFactory.hpp>
#include <com/sun/star/frame/XComponentLoader.hpp>
#include <com/sun/star/beans/PropertyValue.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/util/Color.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/container/XNameAccess.hpp>
#include <com/sun/star/container/XNamed.hpp>
#include <com/sun/star/frame/XStorable.hpp>
#include <com/sun/star/text/XText.hpp>
#include <iostream>
using namespace com::sun::star::uno;
using namespace com::sun::star::lang;
using namespace com::sun::star::beans;
using namespace com::sun::star::frame;
using namespace com::sun::star::sheet;
using namespace com::sun::star::container;
using namespace com::sun::star::table;
using namespace com::sun::star::util;
using namespace com::sun::star::text;
using namespace rtl;
using namespace cppu;
using namespace std;
int main() {
//////////////////////////////////////////////////
	Reference<XComponentContext> xContext;
	Reference<XMultiComponentFactory> xMSFactory;
	Reference<XComponent> xDocument;
	// get the remote office component context
	xContext = bootstrap();
	// get the remote office service manager
	xMSFactory = xContext->getServiceManager();
	// get an instance of the remote office desktop UNO service
	// and query the XComponentLoader interface
	Reference<XInterface> desktop = xMSFactory->createInstanceWithContext(
			OUString::createFromAscii("com.sun.star.frame.Desktop"), xContext);
	Reference<XComponentLoader> rComponentLoader(desktop, UNO_QUERY_THROW);
	// the boolean property Hidden tells the office to open a file in hidden mode
	Sequence<PropertyValue> loadProps(1);
	loadProps[0].Name = OUString::createFromAscii("Hidden");
	loadProps[0].Value = Any(true);	//new Boolean(true);
	//get an instance of the spreadsheet document
	xDocument = rComponentLoader->loadComponentFromURL(
			OUString::createFromAscii("private:factory/scalc"),
			OUString::createFromAscii("_blank"), 0, loadProps);
	//query for a XSpreadsheetDocument interface
	Reference<XSpreadsheetDocument> rSheetDoc(xDocument, UNO_QUERY);
	//use it to get the XSpreadsheets interface
	Reference<XSpreadsheets> rSheets = rSheetDoc->getSheets();
	//query for the XIndexAccess interface
	Reference<XIndexAccess> xSheetsIA(rSheets, UNO_QUERY);
	Any sheet = xSheetsIA->getByIndex(0);
	Reference<XSpreadsheet> rSpSheet(sheet, UNO_QUERY);
	double num = 1.0;
	for (int i = 0; i < 10; i++, num += 2.0) {
		Reference<XCell> cell = rSpSheet->getCellByPosition(i, 0);
		cell->setValue(num);
//        cell->setPropertyValue
	}
	Reference<XCell> cell = rSpSheet->getCellByPosition(10, 0);
//	cell->setFormula(OUString::createFromAscii("=A1+B1"));
	Reference<XText> xx(cell,UNO_QUERY);
	xx->insertString(xx->createTextCursor(),
				OUString::createFromAscii("sfsfsdfsss"), false);
//	cout<<cell->getValue()<<endl;
	cout<<xx->getString()<<endl;
//	cout << string((const char*) cell->getFormula().getStr()) << endl;
//	X aa;
	//////////////////////////
	Reference<XPropertySet> rCellProps(cell, UNO_QUERY);
//	Any PropVal;
//	//PropVal <<= OUString::createFromAscii("Result");
//	PropVal <<= (Color) (0xff0000);
//	rCellProps->setPropertyValue(OUString::createFromAscii("CharColor"),
//			PropVal);
	//////////////////////////
	Reference<XStorable> rStore(xDocument, UNO_QUERY);
	Sequence<PropertyValue> storeProps(1);
	storeProps[0].Name = OUString::createFromAscii("FilterName");
	storeProps[0].Value = Any(
			OUString::createFromAscii("Calc MS Excel 2007 XML"));
	rStore->storeAsURL(OUString::createFromAscii("file:///tmp/MyTest2.xlsx"),
			storeProps);
	xDocument->dispose();
	return 0;
}
#endif
