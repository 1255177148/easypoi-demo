package com.zhan.easypoi.demo.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.rtf.RtfWriter2;
import com.zhan.easypoi.demo.entity.ExportDataEntity;
import com.zhan.easypoi.demo.mapper.ExportDataMapper;
import com.zhan.easypoi.demo.service.ExportDataService;
import com.zhan.easypoi.demo.util.BarCodeUtil;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Author elvis
 * @Date 2021/7/9 16:26
 */
@Service
public class ExportDataServiceImpl extends ServiceImpl<ExportDataMapper, ExportDataEntity> implements ExportDataService {

    @Override
    public void exportWordWithImage() throws IOException, DocumentException {
        List<ExportDataEntity> exportDataEntityList = list();
        List<String> barcodes = exportDataEntityList.stream().map(ExportDataEntity::getBarcodeNumber).collect(Collectors.toList());
        String fileName = "d:/img2doc.doc";
        Document doc = new Document(PageSize.A4);
        doc.setMargins(20,20,20,20);
        /**
         * 建立一个书写器与document对象关联,通过书写器可以将文档写入到输出流中
         */
        RtfWriter2.getInstance(doc, new FileOutputStream(fileName));
        doc.open();
        for (String barcode : barcodes){
            Image image = Image.getInstance(imageToBytes(BarCodeUtil.getBarCode(barcode)));
            image.setSpacingAfter(20);
            image.setSpacingBefore(20);
            doc.add(image);
            Paragraph paragraph = new Paragraph("\n");
            doc.add(paragraph);
        }
        doc.close();
    }


    /**
     * BufferedImage转byte[]
     *
     * @param bImage BufferedImage对象
     * @return byte[]
     * @auth zhy
     */
    private static byte[] imageToBytes(BufferedImage bImage) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            ImageIO.write(bImage, "jpg", out);
        } catch (IOException e) {
            //log.error(e.getMessage());
        }
        return out.toByteArray();
    }
}
