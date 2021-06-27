package com.zhan.easypoi.demo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import java.util.Date;

/**
 * @Author elvis
 * @Date 2021/6/27 14:23
 */
@Data
public class StudentEntity {

    private String id;

    @Excel(name = "学生姓名", height = 20, width = 30)
    private String name;

    @Excel(name = "学生性别", replace = {"女_0", "男_1"})
    private int sex;

    @Excel(name = "出生日期", format = "yyyy-MM-dd", width = 20)
    private Date birthday;

    @Excel(name = "进校日期", format = "yyyy-MM-dd")
    private Date registrationDate;
}
