#ifndef EXCELPOI_HPP
#define EXCELPOI_HPP

#include <iostream>
#include <jni.h>
using namespace std;

#define USE_RINTERNALS
#define MAX_EXCEL_COLS 256

extern "C" {

#include <R.h>
#include <Rinternals.h>
#include <Rdefines.h>
#include <R_ext/Rdynload.h>

  void R_init_Rexcelpoi(DllInfo *info);
  void R_unload_Rexcelpoi(DllInfo *info);

  int JNI_init(void);

  SEXP R2xls(SEXP r_object,
             SEXP filename,
             SEXP workbookName);
}

bool writeWB_from_List(SEXP r_object, const char* fname);
bool writeWB_from_Dataframe(SEXP r_object, const char* fname, const char* workbookName);
bool writeWB_from_Matrix(SEXP r_object, const char* fname, const char* workbookName);
bool writeWB_from_Vector(SEXP r_object, const char* fname, const char* workbookName);
bool writeWB_from_Rseries(SEXP r_object, const char* fname, const char* workbookName);

int writeSHEET_from_List(jobject wb, jobject ws, int startrow, SEXP r_object);
int writeSHEET_from_Dataframe(jobject wb, jobject ws, int startrow, SEXP r_object, bool freeze_panes);
int writeSHEET_from_Matrix(jobject wb, jobject ws,  int startrow, SEXP r_object, bool freeze_panes);
int writeSHEET_from_Vector(jobject wb, jobject ws, int startrow, SEXP r_object, bool freeze_panes);
int writeSHEET_from_Rseries(jobject wb, jobject ws, int startrow, SEXP r_object, bool freeze_panes);

bool do_output(const char* fname, jobject wb);
jobject makeWorkbook();
jobject createExcelSheet(jobject WorkBook, const char* sheetName);
void freeze(jobject sheet, int writeColNames, int writeRowNames);
void writeRowNames(jobject sheet, const vector<string>& rowNames, int writeColNms);
void writeColNames(jobject sheet, const vector<string>& colNames, int writeRowNms);
void writeDates2RowNames(jobject workbook, jobject sheet, SEXP dates, int writeColNms);
jstring make_jstring(const char* str);

void writeCell(jobject cell, double value);
void writeCell(jobject cell, int value);
void writeCell(jobject cell, const char* value);
void writeCase(jobject cell, SEXP r_data, int index);

jobject createStyleAlignRight(jobject wb);
jobject createStyleAlignLeft(jobject wb);

void writeAlongRow(jobject sheet, jobject style, SEXP r_object, int index_start, int index_end, int row_num, int start_col);
void writeAlongCol(jobject sheet, jobject style, SEXP r_object, int index_start, int index_end, int col_num, int start_row);

void writeAlongRow(jobject sheet, jobject style, SEXP r_object, int row_num, int start_col);
void writeAlongCol(jobject sheet, jobject style, SEXP r_object, int col_num, int start_row);

void writeAlongRow(jobject sheet, jobject style, const vector<string>& svec, int row_num, int start_col);
void writeAlongCol(jobject sheet, jobject style, const vector<string>& svec, int row_num, int start_col);

void writeAlongRow(jobject sheet, jobject style, const char* ch, int row_num, int start_col);

void ExceptionCheck();

#endif // EXCELPOI_HPP
