package com.wz.excel.write;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.wz.excel.model.ExcelMetaData;

import java.util.List;

/**
 * @user wangzhen
 * @Date 2018/6/25/025 10:18
 */
public class RimlessDefaultWriteHandler<T> extends AbstractWriteHandler<T> {

	public RimlessDefaultWriteHandler() {
		super(new SXSSFWorkbook(1000));
	}

	public RimlessDefaultWriteHandler(Workbook workbook) {
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

		setAllBorder(style);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setFont(font);

		return style;
	}

	protected void setAllBorder(CellStyle cellStyle) {
		// cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
		// cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
		// cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
		// cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
	}

	@Override
	public CellStyle generateCellStyle(ExcelMetaData metaData) {

		CellStyle cellStyle = super.wb.createCellStyle();

		cellStyle.setAlignment(metaData.getExcelCell().align());// 水平对齐方式
		cellStyle.setVerticalAlignment(metaData.getExcelCell().verticalAlgin());// 垂直对齐方式

		// 时间格式
		if (StringUtils.isNotEmpty(metaData.getExcelCell().dateFormat())) {
			cellStyle.setDataFormat(super.wb.createDataFormat().getFormat(metaData.getExcelCell().dateFormat()));
		}

		if (StringUtils.isNotEmpty(metaData.getExcelCell().numberFormat())) {
			cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(metaData.getExcelCell().numberFormat()));
		}

		super.styleMap.put(metaData.getField(), cellStyle);

		return cellStyle;

	}

}
