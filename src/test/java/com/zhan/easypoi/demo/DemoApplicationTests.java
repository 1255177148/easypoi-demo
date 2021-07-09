package com.zhan.easypoi.demo;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.zhan.easypoi.demo.entity.CourseEntity;
import com.zhan.easypoi.demo.entity.StudentEntity;
import com.zhan.easypoi.demo.entity.TeacherEntity;
import com.zhan.easypoi.demo.util.ExcelStyleUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SpringBootTest
class DemoApplicationTests {

    /**
     * 测试根据实体类导出excel
     */
    @Test
    void testExportExcelByEntityClass() throws IOException {

        List<CourseEntity> courseEntities = new ArrayList<>();
        CourseEntity courseEntity = new CourseEntity();
        courseEntity.setId("1");
        courseEntity.setName("测试课程");
        TeacherEntity teacherEntity = new TeacherEntity();
        teacherEntity.setName("张老师");
        teacherEntity.setSex(0);
        courseEntity.setMathTeacher(teacherEntity);

        List<StudentEntity> studentEntities = new ArrayList<>();
        for (int i = 0;i<=1;i++){
            StudentEntity studentEntity = new StudentEntity();
            studentEntity.setName("学生" + i);
            studentEntity.setSex(i);
            studentEntity.setBirthday(new Date());
            studentEntities.add(studentEntity);
        }
        courseEntity.setStudents(studentEntities);
        courseEntities.add(courseEntity);
        Long start = System.currentTimeMillis();
        ExportParams exportParams = new ExportParams("导出测试\n课程", "测试");
        exportParams.setStyle(ExcelStyleUtil.class);
        exportParams.setTitleHeight((short)20);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, CourseEntity.class, courseEntities);
        Long end = System.currentTimeMillis();
        System.out.println("耗时：" + (end-start) + "ms");
        File file = new File("D:/excel/");
        if (!file.exists()){
            file.mkdirs();
        }
        FileOutputStream fileOutputStream = new FileOutputStream("D:/excel/课程导出测试.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

}
