/*
 * oogo_c.h
 *
 *  Created on: Sep 28, 2014
 *      Author: cny
 */

#ifndef OOGO_C_H_
#define OOGO_C_H_
//
#ifdef LIBOOGO_EXPORTS
#define OOGO_EXPORT _declspec(dllexport)
#else
#define OOGO_EXPORT
#endif
//
#ifdef __cplusplus
extern "C" {
#endif
//
#define E_BUF_LEN 2048;
//
typedef struct {
	int code;
	void* calc;
} calc_c;
typedef struct {
	int code;
	void* sheet;
} sheet_c;
//
OOGO_EXPORT int oogo_init();
OOGO_EXPORT int oogo_destory();
OOGO_EXPORT void oogo_cpy_ebuf(char* buf);
//
OOGO_EXPORT calc_c oogo_new_calc();
OOGO_EXPORT calc_c oogo_open_calc(const char* url);
OOGO_EXPORT int oogo_store_calc(calc_c c, const char* filter, const char* url);
OOGO_EXPORT int oogo_close_calc(calc_c c);
//
OOGO_EXPORT sheet_c oogo_sheet_new(calc_c c, const char* name, int idx);
OOGO_EXPORT sheet_c oogo_sheet_i(calc_c c, int idx);
OOGO_EXPORT sheet_c oogo_sheet_n(calc_c c, const char* name);
//
OOGO_EXPORT int oogo_sheet_end_r_l(sheet_c s, int* col, int* row);
//
OOGO_EXPORT int oogo_sheet_set_v(sheet_c s, int x, int y, double num);
OOGO_EXPORT int oogo_sheet_get_v(sheet_c s, int x, int y, double* num);
//
OOGO_EXPORT int oogo_sheet_set_formula(sheet_c s, int x, int y,
		const char* val);
OOGO_EXPORT int oogo_sheet_get_formula_l(sheet_c s, int x, int y, int* l);
OOGO_EXPORT int oogo_sheet_cpy_formula(sheet_c s, int x, int y, char* buf,
		int len);
//
OOGO_EXPORT int oogo_sheet_set_text(sheet_c s, int x, int y, const char* text);
OOGO_EXPORT int oogo_sheet_get_text_l(sheet_c s, int x, int y, int* l);
OOGO_EXPORT int oogo_sheet_cpy_text(sheet_c s, int x, int y, char* buf,
		int len);
//
#ifdef __cplusplus
}
#endif
#endif /* OOGO_C_H_ */
