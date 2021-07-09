package com.zhan.easypoi.demo.service;

import com.baomidou.mybatisplus.extension.service.IService;
import com.lowagie.text.DocumentException;
import com.zhan.easypoi.demo.entity.ExportDataEntity;

import java.io.IOException;

/**
 * @Author elvis
 * @Date 2021/7/9 16:25
 */
public interface ExportDataService extends IService<ExportDataEntity> {

    /**
     * 将图片插入到word并导出
     */
    void exportWordWithImage() throws IOException, DocumentException;
}
