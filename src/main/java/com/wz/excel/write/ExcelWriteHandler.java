package com.wz.excel.write;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;


public interface ExcelWriteHandler<T> {
  
  String DEFAULT_SHEET = "sheet1";
  
  default Workbook write(List<T> list) {
    return write(DEFAULT_SHEET, list);
  }
  
  Workbook write(String sheetName, List<T> list);
  
}
