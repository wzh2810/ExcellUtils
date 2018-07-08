package com.wz.excel.write;

import com.google.common.collect.Lists;
import com.wz.excel.entity.Person;
import com.wz.excel.exception.ExcelException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.List;

public class DefaultWriteHandler<T> extends AbstractWriteHandler<T> {

	public DefaultWriteHandler() {
		super(new SXSSFWorkbook(1000));
	}

	public DefaultWriteHandler(Workbook workbook) {
		super(workbook);
	}

	@Override
	protected int writeHeader(Workbook wb, Sheet sheet, T t, int titleLineNum) {
		Row row = sheet.createRow(titleLineNum);
		row.setHeight((short) 0x180);

		int colNum = 0;
		CellStyle cellStyle = createHeadCellStyle(wb);
		List<String> cellTitles = excelCellUtils.getOrderedCellTitle(t);

		for (String title : cellTitles) {
			// sheet.autoSizeColumn(colNum);

			Cell cell = row.createCell(colNum++);
			cell.setCellValue(title);
			cell.setCellStyle(cellStyle);
		}

		return titleLineNum;
	}

	/**
	 * 创建标题Style
	 *
	 * @param wb
	 * @return
	 */
	private CellStyle createHeadCellStyle(Workbook wb) {
		// 生成一个字体
		Font font = wb.createFont();
		font.setColor(IndexedColors.BLACK.index);
		// font.setFontHeightInPoints((short) 12);
		// font.setBold(true);

		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setAllBorder(style);

		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setFont(font);

		return style;
	}

	public static void main(String[] args) {
		 List<Person> personList = Lists.newArrayList();
	        for (int i = 0; i <  10; i++) {
	            personList.add(Person.getDemoPerson());
	        }

	        ExcelWriteHandler<Person> handler = new DefaultWriteHandler<>();
	        Workbook wb = handler.write(personList);
	        System.out.println("args = [" + wb + "]");
	        String filePath = "D:\\study\\zip\\excel2007-1001w.xlsx";

	        try {
	            OutputStream outputStream = new FileOutputStream(new File(filePath));
	            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(outputStream);
	            wb.write(bufferedOutputStream);
	            wb.close();
	            outputStream.close();
	        } catch (FileNotFoundException e) {
	            throw new ExcelException("write excel exception, file can not found", e);
	        } catch (IOException e) {
	            throw new ExcelException("write excel exception, IOException", e);
	        }

	}

}
