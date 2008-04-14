#ifndef UTIL_HPP
#define UTIL_HPP

#include <string>
#include <vector>
#include <set>
#include <iostream>

extern "C" {
  #include <R.h>
  #include <Rinternals.h>
}

using namespace std;

double scalarReal(const SEXP x);
int scalarInt(const SEXP x);
double getNumericValue(const SEXP x);
double getComFactor(const SEXP x);
string getRTicker(const SEXP x);
string scalarString(const SEXP x);
void setListnames(const SEXP x, std::vector<string> s);
void setColnames(const SEXP x, std::vector<string> s);
void setColnamesDF(const SEXP x, std::vector<string> s);
void setColnamesMatrix(const SEXP x, std::vector<string> s);
void setFakeRowNamesDF(const SEXP x);
bool isPOSIXct(const SEXP x);
void setPOSIXctClass(const SEXP x);
void setDFClass(const SEXP x);
SEXP allocDF(std::vector<SEXPTYPE> colTypes, const int nr);
void setDFelement(const SEXP df, const int c, const int r, const int x);
void setDFelement(const SEXP df, const int c, const int r, const double x);
void setDFelement(const SEXP df, const int c, const int r, const char* s);
std::vector<string> sexp2string(const SEXP str_sexp);
std::vector<string> colnames(const SEXP x);
std::vector<string> rownames(const SEXP x);
std::vector<string> getNames(const SEXP x);
int findDate(const SEXP x, const double& date);
bool isCom(const SEXP x);
bool isTS(const SEXP x);
bool inClass(const SEXP x, const char* str);
void addPOSIXattributes(SEXP x);
void addRseriesClass(SEXP x);
void addDates(SEXP r_object, SEXP r_dates);
int getColumn(const SEXP x, const char* colname);
double* getColPointer(const SEXP x, const int& c);
double* getDates(const SEXP x);
double getDateAt(const SEXP x, const int& i);
void setDates(SEXP x, SEXP dates);
SEXP getDatesSEXP(const SEXP x);
double getMatrixElement(const SEXP x, const int& row, const int& col);
SEXP vector2sexp(std::vector<string> sv);
SEXP set2sexp(set<string> sv);
set<string> sexp2set(const SEXP x);
SEXP getIndexValue(const SEXP object, const int& index);
#endif // UTIL_H

