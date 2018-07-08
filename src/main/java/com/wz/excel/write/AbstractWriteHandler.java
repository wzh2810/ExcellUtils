package com.wz.excel.write;



import com.google.common.base.Strings;
import com.google.common.collect.Maps;
import com.wz.excel.annotation.ExcelCellUtils;
import com.wz.excel.model.ExcelMetaData;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @param <T>
 */
public abstract class AbstractWriteHandler<T> implements ExcelWriteHandler<T> {
  
  protected ExcelCellUtils excelCellUtils = new ExcelCellUtils();
  
  public Workbook wb;
  
  public Map<Field, CellStyle> styleMap = Maps.newConcurrentMap();
  
  
  public AbstractWriteHandler(Workbook wb) {
    this.wb = wb;
  }
  
  @Override
  public Workbook write(String sheetName, List<T> list) {
    
    if (! list.isEmpty()) {
      Sheet sheet = createSheet(this.wb, sheetName);
      
      int titleLineNum = writeHeader(this.wb, sheet, list.get(0), 0);
      
      writeLine(this.wb, sheet, ++ titleLineNum, list);
      
      sheet.setForceFormulaRecalculation(true);
      
      styleMap.clear();
    }
    
    return wb;
  }
  
  
  /**
   * 将对象写入到Sheet中
   *
   * @param wb
   * @param sheet
   * @param startLineNum 从第startLineNum行开始写入
   * @param list
   */
  protected void writeLine(Workbook wb, Sheet sheet, int startLineNum, List<T> list) {
    // CellStyle cellDateStyle = wb.createCellStyle();
    
    if (! CollectionUtils.isEmpty(list)) {
      for (T t : list) {
        Row currentRow = sheet.createRow(startLineNum++);
        List<ExcelMetaData> values = excelCellUtils.getOrderedMedaData(t);
        
        int colNum = 0;
        for (ExcelMetaData metaData : values) {
          //
          Cell cell = currentRow.createCell(colNum++);
          generate(cell, metaData);
          //sheet.setColumnWidth(colNum, 5500);
          //sheet.autoSizeColumn(colNum);
          //TODO 暂时实现 针对公式可进一步抽象
          if (StringUtils.isNotEmpty(metaData.getExcelCell().formula())) {
            cell.setCellType(Cell.CELL_TYPE_FORMULA);
            cell.setCellFormula(metaData.getExcelCell().formula().replace("[rowIndex]", String.valueOf(startLineNum)));
            continue;
          }
          if (metaData.getValue() instanceof Date) {
            cell.setCellValue((Date) metaData.getValue());
          } else if (metaData.getValue() instanceof Boolean) {
            cell.setCellValue((Boolean) metaData.getValue());
          } else if (metaData.getValue() instanceof String) {
            cell.setCellValue((String) metaData.getValue());
          } else if (metaData.getValue() instanceof Double) {
            cell.setCellValue((Double) metaData.getValue());
          } else if (metaData.getValue() instanceof Integer) {
            cell.setCellValue((Integer) metaData.getValue());
          } else if (metaData.getValue() instanceof Short) {
            cell.setCellValue((Short) metaData.getValue());
          } else if (metaData.getValue() instanceof Long) {
            cell.setCellValue((Long) metaData.getValue());
          } else if (metaData.getValue() instanceof Enum) {
            cell.setCellValue(((Enum) metaData.getValue()).name());
          }
        }
      }
    }
  }
  
  
  /**
   * 创建Sheet
   *
   * @param workbook
   * @param sheetName
   * @return 返回创建的SheetName
   */
  protected Sheet createSheet(Workbook workbook, String sheetName) {
    Sheet sheet = null;
    
    if (Strings.isNullOrEmpty(sheetName)) {
      sheet = workbook.createSheet();
    } else {
      if (workbook.getSheet(sheetName) == null) {
        sheet = workbook.createSheet(sheetName);
      } else {
        int num = 1;
        while (workbook.getSheet(sheetName + "-" + num) != null) {
          num++;
        }
        sheet = workbook.createSheet(sheetName + "-" + num);
      }
    }
    return sheet;
  }
  
  /**
   * 获取并生成注解中的样式(可考虑放在子类)
   *
   * @param cell
   * @param metaData
   * @return
   */
  private void generate(Cell cell, ExcelMetaData metaData) {
    CellStyle cellStyle = styleMap.get(metaData.getField()) == null ? generateCellStyle(metaData) : styleMap.get(metaData.getField());
    cell.getSheet().setColumnWidth(cell.getColumnIndex(), metaData.getExcelCell().width());
    //cell.getRow().setHeight(metaData.getExcelCell().rowHeight());
    cell.setCellStyle(cellStyle);
  }
  
  public CellStyle generateCellStyle(ExcelMetaData metaData) {
    CellStyle cellStyle = wb.createCellStyle();
    
    cellStyle.setFillForegroundColor(metaData.getExcelCell().backgroudColor());
    cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
    //全边框
    setAllBorder(cellStyle);
    cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
    cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
    cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
    cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
    cellStyle.setAlignment(metaData.getExcelCell().align());//水平对齐方式
    cellStyle.setVerticalAlignment(metaData.getExcelCell().verticalAlgin());//垂直对齐方式
    cellStyle.setWrapText(metaData.getExcelCell().autoWrap()); //是否自动换行
    
    //时间格式
    if (StringUtils.isNotEmpty(metaData.getExcelCell().dateFormat())) {
      cellStyle.setDataFormat(wb.createDataFormat().getFormat(metaData.getExcelCell().dateFormat()));
    }
    
    if (StringUtils.isNotEmpty(metaData.getExcelCell().numberFormat())) {
      cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(metaData.getExcelCell().numberFormat()));
    }
    
    styleMap.put(metaData.getField(), cellStyle);
    
    return cellStyle;
  }
  
  /**
   * 设置全边框
   *
   * @param cellStyle
   */
  protected void setAllBorder(CellStyle cellStyle) {
    cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
    cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
    cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
    cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
  }
  
  /**
   * 向Sheet中写入标题，写入到第titleLineNum行
   *
   * @param wb
   * @param sheet
   * @param t
   * @param titleLineNum
   * @return
   */
  protected abstract int writeHeader(Workbook wb, Sheet sheet, T t, int titleLineNum);
  
}
