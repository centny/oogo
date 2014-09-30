/*
 * Oogo.h
 *
 *  Created on: Sep 28, 2014
 *      Author: cny
 */

#ifndef OOGO_H_
#define OOGO_H_
#include <cppuhelper/bootstrap.hxx>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/lang/XMultiComponentFactory.hpp>
#include <com/sun/star/frame/XComponentLoader.hpp>
#include <com/sun/star/beans/PropertyValue.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/util/Color.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/table/XColumnRowRange.hpp>
#include <com/sun/star/table/XTableColumns.hpp>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/container/XNameAccess.hpp>
#include <com/sun/star/container/XNamed.hpp>
#include <com/sun/star/frame/XStorable.hpp>
#include <com/sun/star/sheet/XUsedAreaCursor.hpp>
#include <com/sun/star/sheet/XCellRangeAddressable.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include <com/sun/star/text/XText.hpp>
#include <string>
#include <set>
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
namespace oogo {
class Sheet;
class Calc;
class Bootstrap {
private:
	Reference<XComponentContext> context;
	Reference<XMultiComponentFactory> factory;
	Reference<XInterface> desktop;
	Reference<XComponentLoader> loader;
public:
	static string NEW_CALC;
public:
	Bootstrap();
	virtual ~Bootstrap();
	//
	Reference<XComponent> loadComponent(string name);
	Calc* loadCalc(string name);
};
class Calc {
private:
	Reference<XComponent> doc;
	Reference<XSpreadsheetDocument> calc;
	set<Sheet*> sss;
public:
	Calc(Reference<XComponent> doc, Reference<XSpreadsheetDocument> calc);
	virtual ~Calc();
	Sheet* newSheet(string name, int idx);
	Sheet* sheet_i(int idx);
	Sheet* sheet_n(string name);
	void store(string filter, string url);
};
class Sheet {
private:
	Reference<XSpreadsheet> sheet;
public:
	Sheet(Reference<XSpreadsheet> sheet);
	virtual ~Sheet();
	//
	void end_r_l(int* col, int* row);
	//
	void setValue(int x, int y, double num);
	double getValue(int x, int y);
	//
	void setFormula(int x, int y, string val);
	int getFormulaLen(int x, int y);
	string getFormula(int x, int y);
	void copyFormula(int x, int y, char* buf, int len);
	//
	void setText(int x, int y, string text);
	int getTextLen(int x, int y);
	string getText(int x, int y);
	void copyText(int x, int y, char* buf, int len);
};
OUString StringToOUString(string data);
} /* namespace oogo */

#endif /* OOGO_H_ */
