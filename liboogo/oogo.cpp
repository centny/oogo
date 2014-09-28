/*
 * Oogo.cpp
 *
 *  Created on: Sep 28, 2014
 *      Author: cny
 */

#include "../oogo.h"

namespace oogo {
string Bootstrap::NEW_CALC = "private:factory/scalc";
Bootstrap::Bootstrap() {
	this->context = bootstrap();
	// get the remote office service manager
	this->factory = this->context->getServiceManager();
	// get an instance of the remote office desktop UNO service
	// and query the XComponentLoader interface
	Reference<XInterface> desktop = this->factory->createInstanceWithContext(
			OUString::createFromAscii("com.sun.star.frame.Desktop"),
			this->context);
	this->loader = Reference<XComponentLoader>(desktop, UNO_QUERY_THROW);
}
Bootstrap::~Bootstrap() {
}
Reference<XComponent> Bootstrap::loadComponent(string name) {
	//"private:factory/scalc"
	Sequence<PropertyValue> loadProps(1);
	loadProps[0].Name = OUString::createFromAscii("Hidden");
	loadProps[0].Value = Any(true);	//new Boolean(true);
	return this->loader->loadComponentFromURL(StringToOUString(name.c_str()),
			OUString::createFromAscii("_blank"), 0, loadProps);
}
Calc* Bootstrap::loadCalc(string name) {
	Reference<XComponent> doc = this->loadComponent(name);
	return new Calc(doc, Reference<XSpreadsheetDocument>(doc, UNO_QUERY));
}
/////////////////////////////////////////////////////////////////////////////
Calc::Calc(Reference<XComponent> doc, Reference<XSpreadsheetDocument> calc) {
	this->doc = doc;
	this->calc = calc;
}
Calc::~Calc() {
	set<Sheet*>::iterator it = this->sss.begin();
	set<Sheet*>::iterator end = this->sss.end();
	while (it != end) {
		delete *it;
		it++;
	}
	this->sss.clear();
	this->doc->dispose();
}
Sheet* Calc::newSheet(string name, int idx) {
	this->calc->getSheets()->insertNewByName(StringToOUString(name.c_str()),
			idx);
	return this->sheet_n(name);
}
Sheet* Calc::sheet_i(int idx) {
	Reference<XSpreadsheets> sheets = this->calc->getSheets();
	Reference<XIndexAccess> xia(sheets, UNO_QUERY);
	Any sheet = xia->getByIndex(idx);
	Sheet* ss = new Sheet(Reference<XSpreadsheet>(sheet, UNO_QUERY));
	this->sss.insert(ss);
	return ss;
}
Sheet* Calc::sheet_n(string name) {
	Reference<XSpreadsheets> sheets = this->calc->getSheets();
	Reference<XNameAccess> xia(sheets, UNO_QUERY);
	Any sheet = xia->getByName(StringToOUString(name.c_str()));
	Sheet* ss = new Sheet(Reference<XSpreadsheet>(sheet, UNO_QUERY));
	this->sss.insert(ss);
	return ss;
}
void Calc::store(string filter, string url) {
	Reference<XStorable> rStore(this->doc, UNO_QUERY);
	Sequence<PropertyValue> storeProps(1);
	storeProps[0].Name = OUString::createFromAscii("FilterName");
	storeProps[0].Value = Any(StringToOUString(filter.c_str()));
	rStore->storeAsURL(StringToOUString(url.c_str()), storeProps);
}
Sheet::Sheet(Reference<XSpreadsheet> sheet) {
	this->sheet = sheet;
}
Sheet::~Sheet() {

}
void Sheet::setValue(int x, int y, double num) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	cell->setValue(num);
}
double Sheet::getValue(int x, int y) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	return cell->getValue();
}
void Sheet::setFormula(int x, int y, string val) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	cell->setFormula(StringToOUString(val));
}
string Sheet::getFormula(int x, int y) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	OString os = OUStringToOString(cell->getFormula(), RTL_TEXTENCODING_UTF8);
	return string(os.getStr());
}
void Sheet::copyFormula(int x, int y, char* buf, int len) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	OString os = OUStringToOString(cell->getFormula(), RTL_TEXTENCODING_UTF8);
	strncpy(buf, os.getStr(), len);
}
int Sheet::getFormulaLen(int x, int y) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	OString os = OUStringToOString(cell->getFormula(), RTL_TEXTENCODING_UTF8);
	return os.getLength();
}
void Sheet::setText(int x, int y, string text) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	Reference<XText> xt(cell, UNO_QUERY);
	xt->insertString(xt->createTextCursor(), StringToOUString(text), false);
}
int Sheet::getTextLen(int x, int y) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	Reference<XText> xt(cell, UNO_QUERY);
	OString os = OUStringToOString(xt->getString(), RTL_TEXTENCODING_UTF8);
	return os.getLength();
}
string Sheet::getText(int x, int y) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	Reference<XText> xt(cell, UNO_QUERY);
	OString os = OUStringToOString(xt->getString(), RTL_TEXTENCODING_UTF8);
	return string(os.getStr());
}
void Sheet::copyText(int x, int y, char* buf, int len) {
	Reference<XCell> cell = sheet->getCellByPosition(x, y);
	Reference<XText> xt(cell, UNO_QUERY);
//	const char* dest = (const char*) xt->getString().getStr();
	OString os = OUStringToOString(xt->getString(), RTL_TEXTENCODING_UTF8);
	strncpy(buf, os.getStr(), len);
}
OUString StringToOUString(string data) {
	return OStringToOUString(OString(data.c_str(), data.size()),
	RTL_TEXTENCODING_UTF8);
}
} /* namespace oogo */
