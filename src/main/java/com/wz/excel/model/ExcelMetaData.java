package com.wz.excel.model;

import java.lang.reflect.Field;

import com.wz.excel.annotation.ExcelCell;

public class ExcelMetaData {

	private ExcelCell excelCell;

	private Field field;

	private Object value;

	public ExcelCell getExcelCell() {
		return excelCell;
	}

	public void setExcelCell(ExcelCell excelCell) {
		this.excelCell = excelCell;
	}

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}
}
