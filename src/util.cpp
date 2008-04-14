#include <cstddef>  // For size_t
#include "util.hpp"

double scalarReal(const SEXP x) {
  if(x==R_NilValue||TYPEOF(x)!=REALSXP) {
    return NA_REAL;
  }

  return REAL(x)[0];
}

int scalarInt(const SEXP x) {
  if(x==R_NilValue||TYPEOF(x)!=INTSXP) {
    return NA_INTEGER;
  }

  return INTEGER(x)[0];
}

string scalarString(const SEXP x) {
  string s;

  if(x!=R_NilValue && TYPEOF(x)==STRSXP ) {
    s=CHAR(STRING_ELT(x, 0));
  } else {
    s = "";
  }
  return s;
}


double getNumericValue(const SEXP x) {
  switch(TYPEOF(x)) {
  case REALSXP:
    return scalarReal(x);
    break;
  case INTSXP:
    return (double) scalarInt(x);
    break;
  default:
    return NA_REAL;
    break;
  }
}

void setListnames(const SEXP x, std::vector<string> s) {
  SEXP lnames;

  if(TYPEOF(x)!=VECSXP)
    return;

  int n_size = s.size();
  if(n_size!=length(x)) {
    return;
  }

  PROTECT(lnames = allocVector(STRSXP,n_size));

  for(int i=0; i < n_size; i++) {
    SET_STRING_ELT(lnames, i, mkChar(s[i].c_str()));
  }

  setAttrib(x, R_NamesSymbol, lnames);
  UNPROTECT(1);
}

void setColnamesMatrix(const SEXP x, std::vector<string> s) {

  SEXP dimnames, cnames;

  int cn_size = s.size();
  if(cn_size!=ncols(x))
    return;

  PROTECT(cnames = allocVector(STRSXP,cn_size));

  for(int i=0; i < cn_size; i++) {
    SET_STRING_ELT(cnames, i, mkChar(s[i].c_str()));
  }

  PROTECT(dimnames = allocVector(VECSXP, 2));
  SET_VECTOR_ELT(dimnames, 0, R_NilValue);
  SET_VECTOR_ELT(dimnames, 1, cnames);
  setAttrib(x, R_DimNamesSymbol, dimnames);
  UNPROTECT(2);
}


void setColnamesDF(const SEXP x, std::vector<string> s) {

  SEXP cnames;

  int cn_size = s.size();
  if(cn_size!=length(x)) {
    return;
  }

  PROTECT(cnames = allocVector(STRSXP,cn_size));

  for(int i=0; i < cn_size; i++) {
    SET_STRING_ELT(cnames, i, mkChar(s[i].c_str()));
  }

  setAttrib(x, R_NamesSymbol, cnames);
  UNPROTECT(1);
}

void setColnames(const SEXP x, std::vector<string> s) {

  if(isFrame(x)) {
    setColnamesDF(x,s);
  } else {
    setColnamesMatrix(x,s);
  }
}

void setFakeRowNamesDF(const SEXP x) {

  char buff[10];
  SEXP rnames;

  if(!isFrame(x))
    return;

  int nr = length(VECTOR_ELT(x,0));
  PROTECT(rnames = allocVector(STRSXP,nr));

  for(int i=0; i < nr; i++) {
    // convert int to string
    sprintf(buff,"%d",i+1);
    // convert buff to R string
    SET_STRING_ELT(rnames,i,mkChar(buff));
  }
  setAttrib(x,R_RowNamesSymbol,rnames);
  UNPROTECT(1);
}


void setPOSIXctClass(const SEXP x) {

  SEXP c;

  PROTECT(c = allocVector(STRSXP, 1));
  SET_STRING_ELT(c, 0, mkChar("POSIXct"));
  classgets(x,c);
  UNPROTECT(1);
}

void setDFClass(const SEXP x) {

  SEXP c;

  PROTECT(c = allocVector(STRSXP, 1));
  SET_STRING_ELT(c, 0, mkChar("data.frame"));
  classgets(x,c);
  UNPROTECT(1);
}

// alloc a dataframe based on R types and number of rows
SEXP allocDF(std::vector<SEXPTYPE> colTypes, const int& nr) {
  SEXP ans;

  int nc = colTypes.size();

  PROTECT(ans = allocVector(VECSXP,nc));

  for(int i = 0; i < nc; i++) {
    SET_VECTOR_ELT(ans,i,allocVector(colTypes[i],nr));
  }

  setDFClass(ans);
  setFakeRowNamesDF(ans);
  UNPROTECT(1);
  return ans;
}

// c = column
// r = row
// x = value to put in
// df = dataframe
void setDFelement(const SEXP df, const int& c, const int& r, const int& x) {
  if(TYPEOF(VECTOR_ELT(df,c))!=INTSXP) {
    return;
  }

  INTEGER(VECTOR_ELT(df,c))[r] = x;
}

void setDFelement(const SEXP df, const int& c, const int& r, const double& x) {

  if(TYPEOF(VECTOR_ELT(df,c))!=REALSXP) {
    return;
  }

  REAL(VECTOR_ELT(df,c))[r] = x;
}

void setDFelement(const SEXP df, const int& c, const int& r, const char* s) {

  if(TYPEOF(VECTOR_ELT(df,c))!=STRSXP) {
    return;
  }

  SET_STRING_ELT(VECTOR_ELT(df,c),r,mkChar(s));
}

// convert a vector of R strings to a
// c++ vector of strings
std::vector<string> sexp2string(const SEXP str_sexp) {
  vector<string> ans;

  if(str_sexp==R_NilValue || TYPEOF(str_sexp)!=STRSXP) {
    return ans;
  }

  for(int i=0; i < length(str_sexp); i++) {
    ans.push_back(CHAR(STRING_ELT(str_sexp,i)));
  }
  return ans;
}


std::vector<string> colnames(const SEXP x) {

  SEXP names = R_NilValue;
  bool used_protect = false;

  if(x == R_NilValue) {
    vector<string> ans;
    return ans;
  }

  if(isFrame(x)) {
    PROTECT(names = getAttrib(x,R_NamesSymbol));
    used_protect= true;
  } else {
    SEXP dim_names = getAttrib(x,R_DimNamesSymbol);
    if(dim_names != R_NilValue) {
      PROTECT(names = VECTOR_ELT(dim_names,1));
      used_protect = true;
    }
  }

  if(used_protect) {
    UNPROTECT(1);
  }

  return sexp2string(names);
}

std::vector<string> rownames(const SEXP x) {

  SEXP names = R_NilValue;
  bool used_protect = false;

  if(x == R_NilValue) {
    vector<string> ans;
    return ans;
  }

  if(isFrame(x)) {
    PROTECT(names = getAttrib(x,R_RowNamesSymbol));
    used_protect = true;
  } else {
    SEXP dim_names = getAttrib(x,R_DimNamesSymbol);
    if(dim_names != R_NilValue) {
      PROTECT(names = VECTOR_ELT(dim_names,0));
      used_protect = true;
    }
  }

  if(used_protect) {
    UNPROTECT(1);
  }
  return sexp2string(names);
}


std::vector<string> getNames(const SEXP x) {
  SEXP names;
  vector<string> ans;

  if(x==R_NilValue) {
    return ans;
  }

  PROTECT(names = getAttrib(x,R_NamesSymbol));
  UNPROTECT(1);
  return sexp2string(names);
}


bool isCom(const SEXP x) {
  return inClass(x,"coml");
}

bool isTS(const SEXP x) {
  return inClass(x,"fts");
}

bool inClass(const SEXP x, const char* str) {
  SEXP cls;

  // saftey first
  if(x==R_NilValue) {
    return false;
  }

  // FIXME: NOT PROTECTED
  PROTECT(cls = getAttrib(x,R_ClassSymbol));

  if(cls==R_NilValue) {
    UNPROTECT(1);
    return false;
  }

  for(int i = 0; i < length(cls); i++) {
    if(!strcmp(CHAR(STRING_ELT(cls,i)),str)) {
      UNPROTECT(1);
      return true;
    }
  }

  UNPROTECT(1);
  return false;
}

bool isPOSIXct(const SEXP x) {
  SEXP x_cls;

  if(x==R_NilValue) {
    return false;
  }

  PROTECT(x_cls = getAttrib(x, R_ClassSymbol));

  if(x_cls == R_NilValue) {
    UNPROTECT(1);
    return false;
  }

  if(length(x_cls)==2) {
    if(strcmp("POSIXt",CHAR(STRING_ELT(x_cls,0)))==0 && strcmp("POSIXct",CHAR(STRING_ELT(x_cls,1)))==0 ) {
      UNPROTECT(1);
      return true;
    }
  }

  UNPROTECT(1);
  return false;
}

void addPOSIXattributes(SEXP x) {
  SEXP r_dates_class;

  PROTECT(r_dates_class = allocVector(STRSXP, 2));
  SET_STRING_ELT(r_dates_class, 0, mkChar("POSIXt"));
  SET_STRING_ELT(r_dates_class, 1, mkChar("POSIXct"));
  classgets(x, r_dates_class);
  UNPROTECT(1);
}

void addRseriesClass(SEXP x) {
  SEXP r_tseries_class;
  PROTECT(r_tseries_class = allocVector(STRSXP, 2));
  SET_STRING_ELT(r_tseries_class, 0, mkChar("fts"));
  SET_STRING_ELT(r_tseries_class, 1, mkChar("matrix"));
  classgets(x, r_tseries_class);
  UNPROTECT(1);
}

void addDates(SEXP r_object,SEXP r_dates) {
  if(r_dates==R_NilValue) {
    return;
  }
  setAttrib(r_object,install("dates"),r_dates);
}

int getColumn(const SEXP x, const char* colname) {
  string s(colname);

  vector<string> cnames = colnames(x);

  for(unsigned int i = 0;i < cnames.size(); i++) {
    if(s == cnames[i]) {
      return i;
    }
  }

  return -1;
}

// returns pointer to a vector of floats
// matching the col number
double* getColPointer(const SEXP x, const int& c) {
  //Rprintf("here\n");
  //Rprintf("col: %d\n",c);

  if(x==R_NilValue || TYPEOF(x)!=REALSXP) {
    return 0;
  }

  //Rprintf("nrows(x): %d",nrows(x));

  if(c < 0 || c >= ncols(x) ) {
    Rprintf("column outside of bounds.\n");
    return 0;
  }

  double *xptr = REAL(x);
  return &(xptr[nrows(x)*c]);
}

double *getDates(const SEXP x) {
  return REAL(getAttrib(x,install("dates")));
}

double getDateAt(const SEXP x, const int& i) {
  // FIXME: out of bounds check
  return getDates(x)[i];
}

void setDates(SEXP x, SEXP dates) {
  setAttrib(x,install("dates"),dates);
}

SEXP getDatesSEXP(const SEXP x) {
  return getAttrib(x,install("dates"));
}

double getMatrixElement(const SEXP x, const int& row, const int& col) {
  if(TYPEOF(x)!=REALSXP) {
    return NA_REAL;
  }

  int nc = ncols(x);
  int nr = nrows(x);

  if(row >= nr || nr < 0 || col >= nc || col < 0) {
    return NA_REAL;
  }

  return REAL(x)[row+col*nr];
}

SEXP vector2sexp(std::vector<string> sv) {
  int sv_size = sv.size();

  SEXP ans = allocVector(STRSXP,sv_size);

  for(int i = 0; i < sv_size; i++) {
    SET_VECTOR_ELT(ans,i,mkChar(sv[i].c_str()));
  }

  return ans;
}

SEXP set2sexp(set<string> sv) {
  int sv_size = sv.size();

  SEXP ans = allocVector(STRSXP,sv_size);
  set<string>::iterator it;
  int i = 0;
  for( it = sv.begin(); it != sv.end(); it++ ) {
    SET_VECTOR_ELT(ans,i,mkChar((*it).c_str()));
    i++;
  }

  return ans;
}

set<string> sexp2set(const SEXP x) {
  set<string> ans;

  if(TYPEOF(x)!=STRSXP) {
    return ans;
  }

  for(int i = 0; i < length(x); i++) {
    ans.insert(string(CHAR(VECTOR_ELT(x,i))));
  }

  return ans;

}

SEXP getIndexValue(const SEXP object, const int& index) {

  SEXPTYPE obj_type = TYPEOF(object);
  SEXP ans = allocVector(obj_type,1);

  switch(obj_type) {
  case REALSXP :
    //Rprintf("REAL: %f",REAL(object)[index]);
    REAL(ans)[0] = REAL(object)[index];
    break;
  case LGLSXP :
  case INTSXP :
    //Rprintf("INT: %d",INTEGER(object)[index]);
    INTEGER(ans)[0] = INTEGER(object)[index];
    break;
  default :
    return R_NilValue;
  }

  return ans;
}
