package com.zhan.easypoi.demo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * @Author elvis
 * @Date 2021/6/27 10:21
 */
@Data
public class TeacherEntity {

    @Excel(name = "教师姓名", height = 20, width = 30)
    private String name;

    @Excel(name = "教师性别", replace = {"女_0","男_1"}, needMerge = true)
    private int sex;
}
