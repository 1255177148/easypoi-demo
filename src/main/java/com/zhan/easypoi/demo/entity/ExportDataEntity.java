package com.zhan.easypoi.demo.entity;

import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

/**
 * @Author elvis
 * @Date 2021/7/9 16:18
 */
@TableName("export_data")
@Data
public class ExportDataEntity {

    @TableId
    private Long id;

    private String barcodeNumber;
}
