package com.zhan.easypoi.demo.controller;

import com.lowagie.text.DocumentException;
import com.zhan.easypoi.demo.service.ExportDataService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * @Author elvis
 * @Date 2021/7/9 17:40
 */
@RestController
public class TestController {

    @Autowired
    private ExportDataService exportDataService;

    @GetMapping("/test/export")
    public void export(HttpServletResponse response) throws IOException, DocumentException {
        exportDataService.exportWordWithImage(response);
    }

    @GetMapping("/test/import")
    public void importData() throws Exception {
        exportDataService.testImport();
    }
}
