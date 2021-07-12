package com.zhan.easypoi.demo.util;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.License;
import com.aspose.words.Paragraph;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * @Author elvis
 * @Date 2021/7/12 9:35
 */
@Slf4j
public class WordToPdfUtil {

    public static void doc2pdf(InputStream fileInput, OutputStream fileOutput) {
        // 验证License
        try {
            // 验证License
            loadLicense();
            Document doc = new Document(fileInput);
            // 去水印
            removeWatermark(doc);
            doc.save(fileOutput, SaveFormat.PDF);
        } catch (Exception e) {
            log.error("doc change to pdf failure: " + e.getMessage());
        }
    }

    /**
     * Inserts a watermark into a document.
     *
     * @param doc           The input document.
     * @param watermarkText Text of the watermark.
     */
    /*private static void insertWatermarkText(Document doc, String watermarkText) throws Exception {
        // Create a watermark shape. This will be a WordArt shape.
        // You are free to try other shape types as watermarks.
        java.awt.Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        // Set up the text of the watermark.
        watermark.getTextPath().setText(watermarkText);
        watermark.getTextPath().setFontFamily("Arial");
        watermark.setWidth(500);
        watermark.setHeight(100);
        // Text will be directed from the bottom-left to the top-right corner.
        watermark.setRotation(-40);
        // Remove the following two lines if you need a solid black text.
        watermark.getFill().setColor(Color.GRAY); // Try LightGray to get more Word-style watermark
        watermark.setStrokeColor(Color.GRAY); // Try LightGray to get more Word-style watermark
        // Place the watermark in the page center.
        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        watermark.setWrapType(WrapType.NONE);
        watermark.setVerticalAlignment(VerticalAlignment.CENTER);
        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // Create a new paragraph and append the watermark to this paragraph.
        Paragraph watermarkPara = new Paragraph(doc);
        watermarkPara.appendChild(watermark);
        // Insert the watermark into all headers of each document section.
        for (Section sect : doc.getSections()) {
            // There could be up to three different headers in each section, since we want
            // the watermark to appear on all pages, insert into all headers.
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
        }
    }*/

    /**
     * 插入水印
     *
     * @param watermarkPara
     * @param sect
     * @param headerType
     * @throws Exception
     */
    private static void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) throws Exception {
        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
        if (header == null) {
            // There is no header of the specified type in the current section, create it.
            header = new HeaderFooter(sect.getDocument(), headerType);
            sect.getHeadersFooters().add(header);
        }

        // Insert a clone of the watermark into the header.
        header.appendChild(watermarkPara.deepClone(true));
    }

    /**
     * 移除全部水印
     *
     * @param doc
     * @throws Exception
     */
    private static void removeWatermark(Document doc) throws Exception {
        for (Section sect : doc.getSections()) {
            // There could be up to three different headers in each section, since we want
            // the watermark to appear on all pages, insert into all headers.
            removeWatermarkFromHeader(sect, HeaderFooterType.HEADER_PRIMARY);
            removeWatermarkFromHeader(sect, HeaderFooterType.HEADER_FIRST);
            removeWatermarkFromHeader(sect, HeaderFooterType.HEADER_EVEN);
        }
    }

    /**
     * 移除指定Section的水印
     *
     * @param sect
     * @param headerType
     * @throws Exception
     */
    private static void removeWatermarkFromHeader(Section sect, int headerType) throws Exception {
        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
        if (header != null) {
            header.removeAllChildren();
        }
    }

    /**
     * 从Classpath（jar文件中）中读取License
     */
    private static void loadLicense() {
        //返回读取指定资源的输入流
        License license = new License();
        try (InputStream is = Thread.currentThread().getContextClassLoader()
                .getResourceAsStream("templates/license.xml")) {
            license.setLicense(is);
        } catch (Exception e) {
            log.warn("load word licenses failure");
        }
    }
}
