package com.wz.excel.entity;

import java.util.Date;
import java.util.Random;

import org.apache.poi.hssf.util.HSSFColor;

import com.wz.excel.annotation.ExcelCell;

public class Person {

	@ExcelCell(priority = "A", cellTitle = "姓名",width = 2000,backgroudColor = HSSFColor.YELLOW.index)
    private String name;

    @ExcelCell(priority = "B", cellTitle = "age",width = 2500)
    private int age;

    @ExcelCell(priority = "C", cellTitle = "age1",width = 3000)
    private int age1;

    @ExcelCell(priority = "D", cellTitle = "性别",width = 3500)
    private Sex sex;

    @ExcelCell(priority = "E", cellTitle = "生日",dateFormat = "yyyy/MM/dd")
    private Date birthDay;

    @ExcelCell(priority = "F", cellTitle = "挂掉日",dateFormat = "yyyy/MM/dd")
    private Date deadDay;

    @ExcelCell(priority = "26",cellTitle = "sum",formula = "SUM(C[rowIndex],B[rowIndex])")
    private int sum;

    @ExcelCell(priority = "27",cellTitle = "shang",formula = "C[rowIndex]/B[rowIndex]",numberFormat = "0.00%")
    private float shang;


    public static Person getDemoPerson() {
        Date date = new Date();
        date.setTime(date.getTime() - new Random().nextInt(10000) * 1000 * 10000);
        Date date2 = new Date();
        date2.setTime(date.getTime() + new Random().nextInt(10000) * 1000 * 10000);
        Person person = new Person("Tom", new Random().nextInt(100), Sex.man, date, new Random().nextInt(100),date2);
        return person;
    }


    public enum Sex {
        man,
        woman;
    }

    public Person() {
    }

    public Person(String name, int age, Sex sex, Date birthDay,int age1,Date deadDay) {
        this.name = name;
        this.age = age;
        this.sex = sex;
        this.birthDay = birthDay;
        this.age1 = age1;
        this.deadDay = deadDay;
    }
}
