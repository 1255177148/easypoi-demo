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
import com.zhan.easypoi.demo.util.WordToPdfUtil;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Author elvis
 * @Date 2021/7/9 16:26
 */
@Service
public class ExportDataServiceImpl extends ServiceImpl<ExportDataMapper, ExportDataEntity> implements ExportDataService {

    @Override
    public void exportWordWithImage(HttpServletResponse response) throws IOException, DocumentException {
        List<ExportDataEntity> exportDataEntityList = list();
        List<String> barcodes = exportDataEntityList.stream().map(ExportDataEntity::getBarcodeNumber).collect(Collectors.toList());
        String fileName = "test";
        Document doc = new Document(PageSize.A4);
        doc.setMargins(0,0,50,0);
        /**
         * 建立一个书写器与document对象关联,通过书写器可以将文档写入到输出流中
         */
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        RtfWriter2.getInstance(doc, outputStream);
        doc.open();
        int total = barcodes.size();
        int number = 1;
        for (String barcode : barcodes){
            BufferedImage bufferedImage = BarCodeUtil.getBarCode(barcode);
            String insertWord = barcode + "\n" + number + "/" + total;
            number++;
            Image image = Image.getInstance(imageToBytes(BarCodeUtil.insertWords(bufferedImage, insertWord)));
            image.setSpacingAfter(20);
            image.setSpacingBefore(20);
            doc.add(image);
            Paragraph paragraph = new Paragraph("\n");
            doc.add(paragraph);
        }
        doc.close();
        writeResponseParameter(response, fileName);
        InputStream in = new ByteArrayInputStream(outputStream.toByteArray());
        WordToPdfUtil.doc2pdf(in, response.getOutputStream());
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

    public static void writeResponseParameter(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        response.reset();
        response.setCharacterEncoding("utf-8");
        response.setContentType("application/pdf;charset=utf-8");
        response.setHeader("Content-disposition", "attachment; filename=".concat(String.valueOf(URLEncoder.encode(fileName.concat(".pdf"), "UTF-8"))));
    }
}
