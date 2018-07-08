package com.wz.excel.annotation;

import com.google.common.collect.Maps;
import org.apache.commons.collections.CollectionUtils;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;


public enum ExcelOrderedCellCache {
    instance;

    private Map<Class, List<ExcelCell>> orderCellsMap = Maps.newConcurrentMap();

    public List<ExcelCell> getOrderedCells(Class clazz) {
        return this.orderCellsMap.get(clazz);
    }

    public void addOrderedCells(Class clazz, List<ExcelCell> excelCells) {
        if (! CollectionUtils.isEmpty(excelCells)) {
            this.orderCellsMap.put(clazz, excelCells);
        }
    }
}




