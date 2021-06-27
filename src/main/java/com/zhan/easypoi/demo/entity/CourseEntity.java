package com.zhan.easypoi.demo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import cn.afterturn.easypoi.excel.annotation.ExcelEntity;
import lombok.Data;

import java.util.List;

/**
 * @Author elvis
 * @Date 2021/6/27 10:18
 */
@Data
public class CourseEntity {

    private String id;

    @Excel(name = "课程名称", orderNum = "1", width = 25, needMerge = true)
    private String name;

    private TeacherEntity chineseTeacher;

    @ExcelEntity(name = "教师")
    private TeacherEntity mathTeacher;

    @ExcelCollection(name = "学生", orderNum = "4")
    private List<StudentEntity> students;
}
