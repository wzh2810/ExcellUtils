package com.wz.excel.annotation;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import java.lang.annotation.*;



@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelCell {
  
  /**
   * 优先级（列顺序） 数字越小优先级越高
   *
   * @return
   */
  String priority() default "0";
  
  /**
   * 表头
   *
   * @return
   */
  String cellTitle() default "";
  
  /**
   * 日期格式 yyyy-MM-dd hh:mm:ss
   *
   * @return
   */
  String dateFormat() default "";
  
  /**
   * 函数表达式  SUM(A[rowIndex],C[rowIndex])
   * <p>
   * A : 列名
   * <p>
   * [rowIndex] 固定格式 占位符（行标）,自动解析为对应数据行的下标 解析后: SUM(A2,C2) ......
   *
   * @return
   */
  String formula() default "";
  
  
  int width() default 5000;
  
  
  short backgroudColor() default HSSFColor.WHITE.index;
  
  /**
   * 数字格式化  整数:0  保留两位小数:0.00 百分比0%  小数百分比0.00%
   *
   * @return
   */
  String numberFormat() default "";
  
  /**
   * 说明行
   *
   * @return
   */
  String cellDesc() default "";

    /**
     * 对齐方式 默认左对齐 左：HSSFCellStyle.ALIGN_LEFT  中:HSSFCellStyle.ALIGN_CENTER 右:HSSFCellStyle.ALIGN_RIGHT
     * @return
     */
    short align() default HSSFCellStyle.ALIGN_LEFT;

    /**
     * 垂直对齐方式 默认区中
     */
    short verticalAlgin() default HSSFCellStyle.VERTICAL_CENTER;

  /**
   * 自动换行 默认开启 true  关闭:false
   * @return
   */
    boolean autoWrap() default true;

    /**
     * 行高
     * @return
     */
    short rowHeight() default 300;

}
