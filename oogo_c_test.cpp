/*
 * oogo_c_test.cpp
 *
 *  Created on: Sep 28, 2014
 *      Author: cny
 */
#ifdef _DEV__
#include "oogo_c.h"
#include "oogo.h"
#include <iostream>
using namespace std;
using namespace oogo;
void testoogo() {
	cout << "init:" << oogo_init() << endl;
	////////////////////////////////
	///
	calc_c calc = oogo_new_calc();
	cout << "new:" << calc.code << endl;
	sheet_c ss = oogo_sheet_i(calc, 0);
	cout << "sheet:" << ss.code << endl;
	for (int i = 0; i < 5; i++) {
		for (int j = 0; j < 10; j++) {
			oogo_sheet_set_v(ss, i, j, i * j * 4);
		}
	}
	cout << "text:" << oogo_sheet_set_text(ss, 6, 0, "abc1") << endl;
	cout << "text:" << oogo_sheet_set_text(ss, 6, 1, "这是中文，chinese") << endl;
	cout << "formula:" << oogo_sheet_set_formula(ss, 6, 3, "=B2+C5") << endl;
	cout << "store:"
			<< oogo_store_calc(calc, "Calc MS Excel 2007 XML",
					"file:///tmp/ttt.xlsx") << endl;
	cout << "close:" << oogo_close_calc(calc) << endl;
	////////////////////////////////
	///
	cout << "<---------Calc--------->" << endl;
	calc = oogo_open_calc("file:///tmp/ttt.xlsx");
	cout << "open:" << calc.code << endl;
	ss = oogo_sheet_i(calc, 0);
	cout << "sheet:" << ss.code << endl;
	double num = 0;
	cout << "get:" << oogo_sheet_get_v(ss, 4, 9, &num) << endl;
	cout << "getv:" << num << endl;
	cout << "get:" << oogo_sheet_get_v(ss, 6, 3, &num) << endl;
	cout << "getv:" << num << endl;
	int l = 0;
	char buf[1024] = { 0 };
	cout << "len:" << oogo_sheet_get_text_l(ss, 6, 0, &l) << endl;
	cout << "lenv:" << l << endl;
	cout << "text:" << oogo_sheet_cpy_text(ss, 6, 0, buf, l) << endl;
	cout << "textv:" << string(buf) << endl;
	cout << "len:" << oogo_sheet_get_text_l(ss, 6, 1, &l) << endl;
	cout << "lenv:" << l << endl;
	cout << "text:" << oogo_sheet_cpy_text(ss, 6, 1, buf, l) << endl;
	cout << "textv:" << string(buf) << endl;
	//
	cout << "close:" << oogo_close_calc(calc) << endl;
	////////////////////////////////
	///
	cout << oogo_destory() << endl;
}
int main() {
	testoogo();	//
//	testnew();
	return 0;
}

#endif

