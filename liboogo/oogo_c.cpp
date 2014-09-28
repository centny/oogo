/*
 * oogo_c.cpp
 *
 *  Created on: Sep 28, 2014
 *      Author: cny
 */
#include "../oogo.h"
#include "../oogo_c.h"
#include <com/sun/star/uno/Exception.hpp>
using namespace oogo;
#ifdef __cplusplus
extern "C" {
#endif
Bootstrap *oogo__ = 0;
char ebuf[2048] = { 0 };
void oogo_cpy_ebuf(char* buf) {
	strcpy(buf, ebuf);
	memset(ebuf, 0, sizeof(ebuf));
}
void oogo_ebuf_c(const char* msg) {
	int len = strlen(msg);
	len = len > 2048 ? 2048 : len;
	strncpy(ebuf, msg, len);
}
void oogo_ebuf(string msg) {
	int len = msg.size();
	len = len > 2048 ? 2048 : len;
	msg.copy(ebuf, 0, len);
}
void oogo_ebuf_s(OUString msg) {
	OString os = OUStringToOString(msg, RTL_TEXTENCODING_UTF8);
	oogo_ebuf(os.getStr());
}
void oogo_ebuf_e(Exception &e) {
	oogo_ebuf_s(e.Message);
}
int oogo_init() {
	try {
		oogo__ = new oogo::Bootstrap();
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
int oogo_destory() {
	if (oogo__) {
		delete oogo__;
		oogo__ = 0;
		return 0;
	} else {
		return 1;
	}
}
//
calc_c oogo_new_calc() {
	calc_c calc;
	calc.calc = 0;
	if (oogo__ == 0) {
		calc.code = -1;
		oogo_ebuf_c("not initial");
		return calc;
	}
	try {
		calc.calc = oogo__->loadCalc(Bootstrap::NEW_CALC);
		calc.code = 0;
		return calc;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		calc.code = -1;
		return calc;
	}
}
calc_c oogo_open_calc(const char* url) {
	calc_c calc;
	if (oogo__ == 0) {
		calc.code = -1;
		oogo_ebuf_c("not initial");
		return calc;
	}
	try {
		calc.calc = oogo__->loadCalc(string(url));
		calc.code = 0;
		return calc;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		calc.code = -1;
		return calc;
	}
}
//
int oogo_store_calc(calc_c c, const char* filter, const char* url) {
	try {
		Calc *calc = (Calc*) c.calc;
		calc->store(string(filter), string(url));
		return 0;
	} catch (com::sun::star::uno::Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
int oogo_close_calc(calc_c c) {
	Calc *calc = (Calc*) c.calc;
	delete calc;
	return 0;
}
sheet_c oogo_sheet_new(calc_c c, const char* name, int idx) {
	sheet_c sheet;
	try {
		Calc *calc = (Calc*) c.calc;
		sheet.sheet = calc->newSheet(string(name), idx);
		sheet.code = 0;
		return sheet;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		sheet.code = -1;
		return sheet;
	}
}
sheet_c oogo_sheet_i(calc_c c, int idx) {
	sheet_c sheet;
	try {
		Calc *calc = (Calc*) c.calc;
		sheet.sheet = calc->sheet_i(idx);
		sheet.code = 0;
		return sheet;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		sheet.code = -1;
		return sheet;
	}
}
sheet_c oogo_sheet_n(calc_c c, const char* name) {
	sheet_c sheet;
	try {
		Calc *calc = (Calc*) c.calc;
		sheet.sheet = calc->sheet_n(name);
		sheet.code = 0;
		return sheet;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		sheet.code = -1;
		return sheet;
	}
}
////
int oogo_sheet_set_v(sheet_c s, int x, int y, double num) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		sheet->setValue(x, y, num);
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
int oogo_sheet_get_v(sheet_c s, int x, int y, double* num) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		*num = sheet->getValue(x, y);
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 0;
	}
}
////
int oogo_sheet_set_formula(sheet_c s, int x, int y, const char* val) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		sheet->setFormula(x, y, string(val));
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
int oogo_sheet_get_formula_l(sheet_c s, int x, int y, int* l) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		*l = sheet->getFormulaLen(x, y);
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
int oogo_sheet_cpy_formula(sheet_c s, int x, int y, char* buf, int len) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		sheet->copyFormula(x, y, buf, len);
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
////
int oogo_sheet_set_text(sheet_c s, int x, int y, const char* text) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		sheet->setText(x, y, string(text));
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
int oogo_sheet_get_text_l(sheet_c s, int x, int y, int* l) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		*l = sheet->getTextLen(x, y);
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}

int oogo_sheet_cpy_text(sheet_c s, int x, int y, char* buf, int len) {
	try {
		Sheet *sheet = (Sheet*) s.sheet;
		sheet->copyText(x, y, buf, len);
		return 0;
	} catch (Exception &e) {
		oogo_ebuf_e(e);
		return 1;
	}
}
//
#ifdef __cplusplus
}
#endif
