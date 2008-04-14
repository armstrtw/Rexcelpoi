#include <vector>
#include <iostream>
#include "excelpoi.hpp"
#include "util.hpp"

// xls cell types
//enum cell_type {xls_numeric, xls_string, xls_formula, xls_blank, xls_boolean, xls_error};


// java env and VM are static
static JNIEnv* java_poi_env;
static JavaVM* java_poi_vm;

// classes and registered routines are static too

// classes we will use
static jclass cls_WorkBook;
static jclass cls_WorkSheet;
static jclass cls_Row;
static jclass cls_Cell;
static jclass cls_DataFormat;
static jclass cls_CellStyle;
static jclass cls_POIFileSystem;
static jclass cls_FileOutputStream;
static jclass cls_FileInputStream;


// methods we will use

// POIFS
static jmethodID initPOIFS;

// workbook methods
static jmethodID initWorkBook;
static jmethodID initWorkBook_fromPOIFS;
static jmethodID createSheet;
static jmethodID writeWB;
static jmethodID getSheet;
static jmethodID getSheetAt;
static jmethodID getNumberOfSheets;
static jmethodID createFreezePane;

// sheet methods;
static jmethodID freezePlanes;
static jmethodID createRow;
static jmethodID getRow;
static jmethodID getLastRowNum;
static jmethodID setColumnWidth;

// row methods;
static jmethodID getFirstCellNum;
static jmethodID getLastCellNum;
static jmethodID createCell;
static jmethodID getCell;

// cell methods;
static jmethodID setCellType;
static jmethodID setCellValue_double;
static jmethodID setCellValue_formula;
static jmethodID setCellValue_string;
static jmethodID setCellStyle;

static jmethodID getCellType;
static jmethodID getCellValue_double;
static jmethodID getCellValue_string;

// cell style methods
static jmethodID createCellStyle;
static jmethodID setAlignment;

// data format methods
static jmethodID createDataFormat;
static jmethodID setDataFormat;
static jmethodID getFormat;
static jmethodID getBuiltinFormat;

// file stream methods;
static jmethodID openFileOutputStream;
static jmethodID closeFileOutputStream;

static jmethodID openFileInputStream;
static jmethodID closeFileInputStream;

// the names of these functions must match 
// the name of the dll: excelpoi.o or .dll
// do not change them lightly b/c then JNI_init
// will not be called (I learned the hard way)
void R_init_Rexcelpoi(DllInfo *info) {
  // start Java JVM
  if(!JNI_init()) {
    Rprintf("problem loading JVM.\n");
    Rprintf("make sure the POI is in your java CLASSPATH.\n");
  }
}

void R_unload_Rexcelpoi(DllInfo *info) {
  // destroy java JVM
  java_poi_vm->DestroyJavaVM();
}

void ExceptionCheck() {
  if(java_poi_env->ExceptionOccurred()) {
    java_poi_env->ExceptionDescribe();
    java_poi_env->ExceptionClear();
  }
}

/* initialize the JVM
   need to read the classpath
   from R via: Sys.getenv()
   at this point and put it
   into options[1] below */
int JNI_init() {
  vector<string> jArgs;

  // java args
  string cp1("-Djava.class.path=");
  string cp2(getenv("CLASSPATH")); 

  jArgs.push_back(cp1+cp2);

  JavaVMInitArgs vm_args;
  JavaVMOption options[jArgs.size()];

  for(int i = 0; i < jArgs.size(); i++) {
    options[i].optionString = const_cast<char*>(jArgs[i].c_str());
  }
  vm_args.version = JNI_VERSION_1_4;
  vm_args.nOptions = jArgs.size();
  vm_args.options = options;
  vm_args.ignoreUnrecognized = JNI_TRUE;

  if(JNI_CreateJavaVM(&java_poi_vm, (void**) &java_poi_env, &vm_args)==0) {

    // register the routines we will use
    cls_WorkBook= java_poi_env->FindClass("org/apache/poi/hssf/usermodel/HSSFWorkbook");
    cls_WorkSheet= java_poi_env->FindClass("org/apache/poi/hssf/usermodel/HSSFSheet");
    cls_Row= java_poi_env->FindClass("org/apache/poi/hssf/usermodel/HSSFRow");
    cls_Cell= java_poi_env->FindClass("org/apache/poi/hssf/usermodel/HSSFCell");		

    cls_CellStyle = java_poi_env->FindClass("org/apache/poi/hssf/usermodel/HSSFCellStyle");
    cls_DataFormat = java_poi_env->FindClass("org/apache/poi/hssf/usermodel/HSSFDataFormat");

    cls_POIFileSystem= java_poi_env->FindClass("org/apache/poi/poifs/filesystem/POIFSFileSystem");		
    cls_FileOutputStream = java_poi_env->FindClass("java/io/FileOutputStream");
    cls_FileInputStream = java_poi_env->FindClass("java/io/FileInputStream");


    // POIFS
    initPOIFS = java_poi_env->GetMethodID(cls_POIFileSystem,"<init>","(Ljava/io/InputStream;)V");

    // workbook methods
    initWorkBook = java_poi_env->GetMethodID(cls_WorkBook,"<init>","()V");
    initWorkBook_fromPOIFS = java_poi_env->GetMethodID(cls_WorkBook,"<init>","(Lorg/apache/poi/poifs/filesystem/POIFSFileSystem;)V");
    createSheet = java_poi_env->GetMethodID(cls_WorkBook,"createSheet","(Ljava/lang/String;)Lorg/apache/poi/hssf/usermodel/HSSFSheet;");
    getNumberOfSheets = java_poi_env->GetMethodID(cls_WorkBook,"getNumberOfSheets","()I");
    getSheet = java_poi_env->GetMethodID(cls_WorkBook,"getSheet","(Ljava/lang/String;)Lorg/apache/poi/hssf/usermodel/HSSFSheet;");
    getSheetAt = java_poi_env->GetMethodID(cls_WorkBook,"getSheetAt","(I)Lorg/apache/poi/hssf/usermodel/HSSFSheet;");
    writeWB = java_poi_env->GetMethodID(cls_WorkBook,"write","(Ljava/io/OutputStream;)V");

    createDataFormat = java_poi_env->GetMethodID(cls_WorkBook,"createDataFormat","()Lorg/apache/poi/hssf/usermodel/HSSFDataFormat;");
    createCellStyle  = java_poi_env->GetMethodID(cls_WorkBook,"createCellStyle","()Lorg/apache/poi/hssf/usermodel/HSSFCellStyle;");
		
    // sheet methods
    createFreezePane = java_poi_env->GetMethodID(cls_WorkSheet,"createFreezePane","(II)V");
    createRow = java_poi_env->GetMethodID(cls_WorkSheet,"createRow","(I)Lorg/apache/poi/hssf/usermodel/HSSFRow;");
    getRow = java_poi_env->GetMethodID(cls_WorkSheet,"getRow","(I)Lorg/apache/poi/hssf/usermodel/HSSFRow;");
    getLastRowNum = java_poi_env->GetMethodID(cls_WorkSheet,"getLastRowNum","()I");
    setColumnWidth = java_poi_env->GetMethodID(cls_WorkSheet,"setColumnWidth","(SS)V");

    // row methods
    getFirstCellNum = java_poi_env->GetMethodID(cls_Row,"getFirstCellNum","()S");
    getLastCellNum = java_poi_env->GetMethodID(cls_Row,"getLastCellNum","()S");
    createCell = java_poi_env->GetMethodID(cls_Row,"createCell","(S)Lorg/apache/poi/hssf/usermodel/HSSFCell;");
    getCell = java_poi_env->GetMethodID(cls_Row,"getCell","(S)Lorg/apache/poi/hssf/usermodel/HSSFCell;");


    // cell methods

    // get type
    getCellType = java_poi_env->GetMethodID(cls_Cell,"getCellType","()I");

    // set values
    setCellValue_double  = java_poi_env->GetMethodID(cls_Cell,"setCellValue","(D)V");
    setCellValue_formula = java_poi_env->GetMethodID(cls_Cell,"setCellFormula","(Ljava/lang/String;)V");
    setCellValue_string  = java_poi_env->GetMethodID(cls_Cell,"setCellValue","(Ljava/lang/String;)V");

    // set style
    setCellStyle = java_poi_env->GetMethodID(cls_Cell,"setCellStyle","(Lorg/apache/poi/hssf/usermodel/HSSFCellStyle;)V");

    // get values
    getCellValue_double  = java_poi_env->GetMethodID(cls_Cell,"getNumericCellValue","()D");
    getCellValue_string  = java_poi_env->GetMethodID(cls_Cell,"getStringCellValue","()Ljava/lang/String;");

    // cell style
    setDataFormat = java_poi_env->GetMethodID(cls_CellStyle,"setDataFormat","(S)V");
    setAlignment = java_poi_env->GetMethodID(cls_CellStyle,"setAlignment","(S)V");

    // data format 
    getFormat = java_poi_env->GetMethodID(cls_DataFormat,"getFormat","(Ljava/lang/String;)S");
    getBuiltinFormat = java_poi_env->GetStaticMethodID(cls_DataFormat,"getBuiltinFormat","(Ljava/lang/String;)S");

    // file stream methods
    openFileOutputStream = java_poi_env->GetMethodID(cls_FileOutputStream,"<init>","(Ljava/lang/String;)V");
    closeFileOutputStream = java_poi_env->GetMethodID(cls_FileOutputStream,"close","()V");

    openFileInputStream = java_poi_env->GetMethodID(cls_FileInputStream,"<init>","(Ljava/lang/String;)V");
    closeFileInputStream = java_poi_env->GetMethodID(cls_FileInputStream,"close","()V");

    return 1;
  } else {
    cerr << "Failed to create VM!" << endl;
    return 0;
  }

}

jstring make_jstring(const char* str) {
  jstring ans = java_poi_env->NewStringUTF(str);
  return ans;
}

void writeCell(jobject cell, double value) {
  if(cell==NULL || ISNAN(value)) {
    return;
  }
  java_poi_env->CallVoidMethod(cell,setCellValue_double,value);
}

void writeCell(jobject cell, int value) {
  if(cell==NULL) {
    return;
  }
  java_poi_env->CallVoidMethod(cell,setCellValue_double,static_cast<double>(value));
}

void writeCell(jobject cell, const char* value) {
  if(cell==NULL || value==NULL) {
    return;
  }

  jstring value_utf = java_poi_env->NewStringUTF(value);
  java_poi_env->CallVoidMethod(cell,setCellValue_string,value_utf);
  java_poi_env->DeleteLocalRef(value_utf);
}

void writeCase(jobject cell, SEXP r_data, int index) {
  const double epoch = -2208970800.0;

  if(r_data==R_NilValue)
    return;
  switch(TYPEOF(r_data)) {
  case REALSXP:
    if(isPOSIXct(r_data)) {
      double excelDate = round((REAL(r_data)[index] - epoch)/86400) + 2.0;      
      writeCell(cell,excelDate);
    } else {
      writeCell(cell,REAL(r_data)[index]);
    }
    break;
  case LGLSXP:
  case INTSXP:
    writeCell(cell,INTEGER(r_data)[index]);
    break;
  case STRSXP:
    writeCell(cell,CHAR(STRING_ELT(r_data,index)));
    break;
  }
}


jobject makeWorkbook() {
  // create new workbook
  jobject wb = java_poi_env->NewObject(cls_WorkBook,initWorkBook);
  return wb;
}

jobject createExcelSheet(jobject wb, const char* sheetName) {

  // FIXME: cut off sheet name length at 31 chars
  if(wb==NULL || !strlen(sheetName) )
    return NULL;

  // change the sheet name in to a jstring
  jstring jworksheetName = make_jstring(sheetName);

  // make a new sheet on the workbook
  jobject sheet = java_poi_env->CallObjectMethod(wb,createSheet,jworksheetName);
  java_poi_env->DeleteLocalRef(jworksheetName);
	
  if(!sheet) {
    Rprintf("createExcelSheet: failed to create sheet.\n");
  }

  return sheet;
}

jobject createStyleAlignRight(jobject wb) {
  jobject style;
  style = java_poi_env->CallObjectMethod(wb,createCellStyle);
  java_poi_env->CallVoidMethod(style,setAlignment,static_cast<short>(3));
  return style;
}

jobject createStyleAlignLeft(jobject wb) {
  jobject style;
  style = java_poi_env->CallObjectMethod(wb,createCellStyle);
  java_poi_env->CallVoidMethod(style,setAlignment,static_cast<short>(1));
  return style;
}

jobject createStyle(jobject wb, const char* fmt, int alignRight) {
  jobject style, data_fmt;
  jstring jdata_fmt_str;
  short data_fmt_code;

  style = NULL;
  
  if(fmt==NULL || strlen(fmt)==0) {
    return style;
  }

  // make our one format
  jdata_fmt_str = make_jstring(fmt);

  // create a blank the style in the workbook
  style = java_poi_env->CallObjectMethod(wb,createCellStyle);
  // create a blank data fmt in the workbook
  data_fmt = java_poi_env->CallObjectMethod(wb,createDataFormat);

  // get the format code from excel (it's a short)
  data_fmt_code = java_poi_env->CallShortMethod(data_fmt,getFormat,jdata_fmt_str);

  // apply the new format code to the style
  java_poi_env->CallVoidMethod(style,setDataFormat,static_cast<short>(data_fmt_code));

  // apply alignment if desired
  if(alignRight) {
    java_poi_env->CallVoidMethod(style,setAlignment,static_cast<short>(3));
  }

  java_poi_env->DeleteLocalRef(jdata_fmt_str);
  java_poi_env->DeleteLocalRef(data_fmt);

  return style;
}


SEXP R2xls(SEXP r_object,
	   SEXP filename,
	   SEXP worksheetName,
	   SEXP formats)
{ 	
  jstring jfilename;
  jobject sheet;
  bool check;
  SEXP ans;

  ans = allocVector(LGLSXP, 1);
  LOGICAL(ans)[0] = FALSE;

  // check to make sure that JVM loads
  if(!java_poi_env) {
    Rprintf("Java VM failed to start on package load.\n");
    return ans;
  }

  if(filename==R_NilValue) {
    Rprintf("Filename is null.\n");
    return ans;
  }

  const char* fname = CHAR(STRING_ELT(filename,0));
  const char* sheetName = CHAR(STRING_ELT(worksheetName,0));

  // write sheet based on object type
  if(isTS(r_object)) {
    check = writeWB_from_Rseries(r_object,fname,sheetName,formats);
  } else if(isFrame(r_object)) {
    check = writeWB_from_Dataframe(r_object,fname,sheetName,formats);
  } else if(isNewList(r_object)) {
    check = writeWB_from_List(r_object,fname,formats);
  } else if(isMatrix(r_object)) {
    check = writeWB_from_Matrix(r_object,fname,sheetName,formats);
  } else if(isVector(r_object)) {
    check = writeWB_from_Vector(r_object,fname,sheetName,formats);
  }

  if(check) {
    LOGICAL(ans)[0] = TRUE;
  }

  return ans;
}

bool writeWB_from_List(SEXP r_object, const char* fname,SEXP formats) {
  SEXP this_element, this_format;
  jobject wb, ws;
  int check;

  wb = makeWorkbook();

  vector<string> nms = getNames(r_object);

  if(nms.empty()) {
    java_poi_env->DeleteLocalRef(wb);
    Rprintf("need names to write a list.\n");
    return false;
  }

  for(int i = 0; i < length(r_object); i++) {

    this_element = VECTOR_ELT(r_object,i);

    if(formats!=R_NilValue && isNewList(formats)) {
      this_format =  VECTOR_ELT(formats, i%length(formats) );
    } else {
      this_format = formats;
    }

    ws = createExcelSheet(wb, nms[i].c_str());

    // write sheet based on object type
    if(isTS(this_element)) {
      check = writeSHEET_from_Rseries(wb,ws,0,this_element,this_format,true);
    } else if(isFrame(this_element)) {
      check = writeSHEET_from_Dataframe(wb,ws,0,this_element,this_format,true);
    } else if(isNewList(this_element)) {
      check = writeSHEET_from_List(wb,ws,0,this_element,this_format);
    } else if(isMatrix(this_element)) {
      check = writeSHEET_from_Matrix(wb,ws,0,this_element,this_format,true);
    } else if(isVector(this_element)) {
      check = writeSHEET_from_Vector(wb,ws,0,this_element,this_format,true);
    }
    java_poi_env->DeleteLocalRef(ws);
    // if there was a problem w/ a sheet
    if(!check) {
      return false;
    }
  }


  if(!do_output(fname,wb)) {
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }

  java_poi_env->DeleteLocalRef(wb);
  return true;
}



bool writeWB_from_Dataframe(SEXP r_object, const char* fname, const char* worksheetName, SEXP formats) {

  jobject wb = makeWorkbook();
  jobject ws = createExcelSheet(wb, worksheetName);
  if(!writeSHEET_from_Dataframe(wb,ws,0,r_object,formats,true)) {
    java_poi_env->DeleteLocalRef(ws);  
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }

  if(!do_output(fname,wb)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }
  java_poi_env->DeleteLocalRef(ws);
  java_poi_env->DeleteLocalRef(wb);
  return true;

}

bool writeWB_from_Matrix(SEXP r_object, const char* fname, const char* worksheetName, SEXP formats) {

  jobject wb = makeWorkbook();
  jobject ws = createExcelSheet(wb, worksheetName);
  if(!writeSHEET_from_Matrix(wb,ws,0,r_object,formats,true)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }

  if(!do_output(fname,wb)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }
  java_poi_env->DeleteLocalRef(ws);
  java_poi_env->DeleteLocalRef(wb);
  return true;
}

bool writeWB_from_Rseries(SEXP r_object, const char* fname, const char* worksheetName, SEXP formats) {

  jobject wb = makeWorkbook();
  jobject ws = createExcelSheet(wb, worksheetName);
  if(!writeSHEET_from_Rseries(wb,ws,0,r_object,formats,true)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }

  if(!do_output(fname,wb)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }
  java_poi_env->DeleteLocalRef(ws);
  java_poi_env->DeleteLocalRef(wb);
  return true;
}

bool writeWB_from_Vector(SEXP r_object, const char* fname, const char* worksheetName, SEXP formats) {

  jobject wb = makeWorkbook();
  jobject ws = createExcelSheet(wb, worksheetName);
  if(!writeSHEET_from_Vector(wb,ws,0,r_object,formats,true)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }

  if(!do_output(fname,wb)) {
    java_poi_env->DeleteLocalRef(ws);
    java_poi_env->DeleteLocalRef(wb);
    return false;
  }
  java_poi_env->DeleteLocalRef(ws);
  java_poi_env->DeleteLocalRef(wb);
  return true;
}

void writeAlongRow(jobject sheet, jobject style, SEXP r_object, int row_num, int start_col) {
  jobject row, cell;
  
  row = java_poi_env->CallObjectMethod(sheet,createRow,row_num);

  for(int i = 0; i < length(r_object); i++, start_col++)  {
    cell = java_poi_env->CallObjectMethod(row,createCell,start_col);
    writeCase(cell, r_object, i);
    if(style) {
      java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    }
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}

void writeAlongRow(jobject sheet, jobject style, SEXP r_object, int index_start, int index_end, int row_num, int start_col) {
  jobject row, cell;
  
  row = java_poi_env->CallObjectMethod(sheet,createRow,row_num);

  for(int i = index_start; i < index_end; i++, start_col++)  {
    cell = java_poi_env->CallObjectMethod(row,createCell,start_col);
    writeCase(cell, r_object, i);
    if(style) {
      java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    }
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}

void writeAlongRow(jobject sheet, jobject style, const char* ch, int row_num, int start_col) {
  jobject row, cell;
  
  row = java_poi_env->CallObjectMethod(sheet,createRow,row_num);
  cell = java_poi_env->CallObjectMethod(row,createCell,start_col);
  writeCell(cell, ch);
  if(style) {
    java_poi_env->CallVoidMethod(cell,setCellStyle,style);
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}

void writeAlongRow(jobject sheet, jobject style, const vector<string>& svec, int row_num, int start_col) {
  jobject row, cell;
  
  row = java_poi_env->CallObjectMethod(sheet,createRow,row_num);

  for(int i = 0; i < svec.size(); i++, start_col++)  {
    cell = java_poi_env->CallObjectMethod(row,createCell,start_col);
    writeCell(cell, svec[i].c_str());
    if(style) {
      java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    }
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}

void writeAlongCol(jobject sheet, jobject style, SEXP r_object, int col_num, int start_row) {

  jobject row, cell;

  for(int i = 0; i < length(r_object); i++, start_row++)  {
    row = java_poi_env->CallObjectMethod(sheet,getRow,start_row);  
    
    if(row==NULL) {
      row = java_poi_env->CallObjectMethod(sheet,createRow,start_row);
    }

    cell = java_poi_env->CallObjectMethod(row,createCell,col_num);
    writeCase(cell, r_object, i);
    if(style) {
      java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    }
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}

void writeAlongCol(jobject sheet, jobject style, SEXP r_object, int index_start, int index_end, int col_num, int start_row) {

  jobject row, cell;

  for(int i = index_start; i < index_end; i++, start_row++)  {
    row = java_poi_env->CallObjectMethod(sheet,getRow,start_row);  
    
    if(row==NULL) {
      row = java_poi_env->CallObjectMethod(sheet,createRow,start_row);
    }

    cell = java_poi_env->CallObjectMethod(row,createCell,col_num);
    writeCase(cell, r_object, i);
    if(style) {
      java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    }
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}

void writeAlongCol(jobject sheet, jobject style, const vector<string>& svec, int col_num, int start_row) {

  jobject row, cell;

  for(int i = 0; i < svec.size(); i++, start_row++)  {
    row = java_poi_env->CallObjectMethod(sheet,getRow,start_row);  
    
    if(row==NULL) {
      row = java_poi_env->CallObjectMethod(sheet,createRow,start_row);
    }

    cell = java_poi_env->CallObjectMethod(row,createCell,col_num);
    writeCell(cell, svec[i].c_str());
    if(style) {
      java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    }
  }
  java_poi_env->DeleteLocalRef(row);
  java_poi_env->DeleteLocalRef(cell);
}



int writeSHEET_from_List(jobject wb, jobject sheet, int startrow, SEXP r_object, SEXP formats) {

  SEXP this_element, this_format;
  jobject row, cell, style;

  vector<string> nms = getNames(r_object);

  if(nms.empty()) {
    java_poi_env->DeleteLocalRef(wb);
    Rprintf("need names to write a list.\n");
    return false;
  }

  for(int i = 0; i < length(r_object); i++) {

    this_element = VECTOR_ELT(r_object,i);

    if(formats!=R_NilValue && isNewList(formats)) {
      this_format =  VECTOR_ELT(formats, i%length(formats) );
    } else {
      this_format = formats;
    }

    // write table name
    row = java_poi_env->CallObjectMethod(sheet,createRow,startrow++);
    cell = java_poi_env->CallObjectMethod(row,createCell,static_cast<short>(0));
    writeCell(cell,nms[i].c_str());
    style = createStyleAlignLeft(wb);
    java_poi_env->CallVoidMethod(cell,setCellStyle,style);
    java_poi_env->DeleteLocalRef(row);
    java_poi_env->DeleteLocalRef(cell);
    java_poi_env->DeleteLocalRef(style);

    // write sheet based on object type
    if(isTS(this_element)) {
      startrow = writeSHEET_from_Rseries(wb,sheet,startrow,this_element,this_format,false);
    } else if(isFrame(this_element)) {
      startrow = writeSHEET_from_Dataframe(wb,sheet,startrow,this_element,this_format, false);
    } else if(isNewList(this_element)) {
      startrow = writeSHEET_from_List(wb,sheet,startrow,this_element,this_format);
    } else if(isMatrix(this_element)) {
      startrow = writeSHEET_from_Matrix(wb,sheet,startrow,this_element,this_format, false);
    } else if(isVector(this_element)) {
      startrow = writeSHEET_from_Vector(wb,sheet,startrow,this_element,this_format, false);
    }
  }

  return startrow + 1;
}

int writeSHEET_from_Dataframe(jobject wb, jobject sheet, int startrow, SEXP r_object, SEXP formats, bool freeze_panes) {

  jobject style;
  short colNum = 0;
  short colNamesDefined = 0;
  short rowNamesDefined = 0;
  int nr, nc;

  if(sheet==NULL || r_object==R_NilValue) {
    return 0;
  }

  vector<string> cnms  = colnames(r_object);
  vector<string> rnms  = rownames(r_object);

  if(cnms.size()) {
    colNamesDefined = 1;
  }

  if(rnms.size()) {
    rowNamesDefined = 1;
  }

  // write colnames
  if(colNamesDefined) {
    style = createStyleAlignRight(wb);
    writeAlongRow(sheet, style, cnms, startrow++, rowNamesDefined);
    java_poi_env->DeleteLocalRef(style);
  }

  // write rownames
  if(rowNamesDefined) {
    style = createStyleAlignLeft(wb);
    writeAlongCol(sheet, style, rnms, colNum++, startrow);
    java_poi_env->DeleteLocalRef(style);
  }

  // write data
  vector<string> fmts = sexp2string(formats);

  nr = length(VECTOR_ELT(r_object,0));
  nc = length(r_object);

  for(int i = 0; i < nc; i++) {
    if(fmts.size()) {
      style = createStyle(wb,fmts[i%fmts.size()].c_str(),1);
    } else {
      style = NULL;
    }

    writeAlongCol(sheet, style, VECTOR_ELT(r_object,i), colNum++, startrow);
    java_poi_env->DeleteLocalRef(style);
  }
  
  if(freeze_panes) {
    java_poi_env->CallVoidMethod(sheet,createFreezePane,rowNamesDefined,colNamesDefined);
  }

  return startrow + nr + 1;
}


int writeSHEET_from_Matrix(jobject wb, jobject sheet, int startrow, SEXP r_object, SEXP formats, bool freeze_panes) {

  jobject style;
  short colNum = 0;
  short colNamesDefined = 0;
  short rowNamesDefined = 0;
  int nr, nc;

  // can't write to void sheet
  if(sheet==NULL || r_object==R_NilValue) {
    return 0;
  }

  vector<string> cnms  = colnames(r_object);
  vector<string> rnms  = rownames(r_object);

  if(cnms.size()) {
    colNamesDefined = 1;
  }

  if(rnms.size()) {
    rowNamesDefined = 1;
  }

  // write colnames
  if(colNamesDefined) {
    style = createStyleAlignRight(wb);
    writeAlongRow(sheet, style, cnms, startrow++, rowNamesDefined);
    java_poi_env->DeleteLocalRef(style);
  }

  // write rownames
  if(rowNamesDefined) {
    style = createStyleAlignLeft(wb);
    writeAlongCol(sheet, style, rnms, colNum++, startrow);
    java_poi_env->DeleteLocalRef(style);
  }

  // write data
  vector<string> fmts = sexp2string(formats);

  nr = nrows(r_object);
  nc = ncols(r_object);

  for(int i = 0; i < nc; i++) {
    if(fmts.size()) {
      style = createStyle(wb,fmts[i%fmts.size()].c_str(),1);
    } else {
      style = NULL;
    }

    writeAlongCol(sheet, style, r_object, i*nr, i*nr+nr, colNum++, startrow);
    java_poi_env->DeleteLocalRef(style);
  }

  if(freeze_panes) {
    java_poi_env->CallVoidMethod(sheet,createFreezePane,rowNamesDefined,colNamesDefined);
  }

  return startrow + nr + 1;
}


int writeSHEET_from_Vector(jobject wb, jobject sheet, int startrow, SEXP r_object, SEXP formats, bool freeze_panes) {

  jobject style;
  short colNum = 0;
  short rowNamesDefined = 0;

  if(sheet==NULL || r_object==R_NilValue) {
    return false;
  }

  vector<string> nms  = getNames(r_object);
  if(nms.size()) {
    rowNamesDefined = 1;
  }

  // write rownames
  if(rowNamesDefined) {
    style = createStyleAlignLeft(wb);
    writeAlongCol(sheet, style, nms, colNum++, startrow);
    java_poi_env->DeleteLocalRef(style);
  }

  // write data
  vector<string> fmts = sexp2string(formats);
  if(fmts.size()) {
    style = createStyle(wb,fmts[0].c_str(),1);
  } else {
    style = NULL;
  }

  writeAlongCol(sheet, style, r_object, colNum, startrow);
  java_poi_env->DeleteLocalRef(style);

  if(freeze_panes) {
    java_poi_env->CallVoidMethod(sheet,createFreezePane,rowNamesDefined,1);
  }

  return startrow + length(r_object) + 1;
}

int writeSHEET_from_Rseries(jobject wb, jobject sheet, int startrow, SEXP r_object, SEXP formats, bool freeze_panes) {

  jobject style;
  SEXP dts;
  short colNum = 0;
  short colNamesDefined = 0;
  int rowNamesDefined = 1;
  int nr, nc;

  // can't write to void sheet
  if(sheet==NULL || r_object==R_NilValue) {
    return 0;
  }

  vector<string> cnms  = colnames(r_object);

  if(cnms.size()) {
    colNamesDefined = 1;
  }

  // write colnames
  if(colNamesDefined) {
    style = createStyleAlignRight(wb);
    writeAlongRow(sheet, style, cnms, startrow++, rowNamesDefined);
    java_poi_env->DeleteLocalRef(style);
  }


  // make 1st column bigger to view dates properly
  java_poi_env->CallVoidMethod(sheet,setColumnWidth,static_cast<short>(0),static_cast<short>(2600));

  // write rownames
  dts = getDatesSEXP(r_object);
  style = createStyle(wb,"yyyy-mm-dd", 0);
  writeAlongCol(sheet, style, dts, colNum++, startrow);
  java_poi_env->DeleteLocalRef(style);

  // write data
  vector<string> fmts = sexp2string(formats);

  nr = nrows(r_object);
  nc = ncols(r_object);

  for(int i = 0; i < nc; i++) {
    if(fmts.size()) {
      style = createStyle(wb,fmts[i%fmts.size()].c_str(),1);
    } else {
      style = NULL;
    }

    writeAlongCol(sheet, style, r_object, i*nr, i*nr+nr, colNum++, startrow);
    java_poi_env->DeleteLocalRef(style);
  }

  if(freeze_panes) {
    java_poi_env->CallVoidMethod(sheet,createFreezePane,rowNamesDefined,colNamesDefined);
  }

  return startrow + nr + 1;
}


bool do_output(const char* fname, jobject wb) {
  jobject fout;
  jstring jfilename;

  // make a new jstring to hold the filename
  jfilename = make_jstring(fname);

  if(!jfilename) {
    java_poi_env->DeleteLocalRef(jfilename);
    error("could not create filename.");
    return false;
  }

  // create the output stream
  fout = java_poi_env->NewObject(cls_FileOutputStream,openFileOutputStream,jfilename);
  ExceptionCheck();
  java_poi_env->DeleteLocalRef(jfilename);

  if(!fout) {
    error("could not create file");
    return false;
  }

  // write the excel file out
  java_poi_env->CallVoidMethod(wb,writeWB,fout);
  ExceptionCheck();

  // close the output stream
  java_poi_env->CallVoidMethod(fout,closeFileOutputStream);
  ExceptionCheck();
  java_poi_env->DeleteLocalRef(fout);

  return true;
}
