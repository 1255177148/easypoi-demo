package com.zhan.easypoi.demo.service.impl;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.alibaba.fastjson.JSON;
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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author elvis
 * @Date 2021/7/9 16:26
 */
@Service
public class ExportDataServiceImpl extends ServiceImpl<ExportDataMapper, ExportDataEntity> implements ExportDataService {

    private static String REGEX_CHINESE = "[\u4e00-\u9fa5]";// 中文正则

    @Override
    public void exportWordWithImage(HttpServletResponse response) throws IOException, DocumentException {
        List<ExportDataEntity> exportDataEntityList = list();
        List<String> barcodes = exportDataEntityList.stream().map(ExportDataEntity::getBarcodeNumber).collect(Collectors.toList());
        String fileName = "test";
        Document doc = new Document(PageSize.A4);
        doc.setMargins(5,0,50,0);
        /**
         * 建立一个书写器与document对象关联,通过书写器可以将文档写入到输出流中
         */
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        RtfWriter2.getInstance(doc, outputStream);
        doc.open();
        int total = barcodes.size();
        int number = 1;
        for (String barcode : barcodes){
            barcode = barcode.replaceAll(REGEX_CHINESE, "");
            BufferedImage bufferedImage = BarCodeUtil.getBarCode(barcode);
            String insertWord = barcode + "\n" + number + "/" + total;
            number++;
            Image image = Image.getInstance(imageToBytes(BarCodeUtil.insertWords(bufferedImage, insertWord)));
//            image.setSpacingAfter(20);
//            image.setSpacingBefore(20);
            doc.add(image);
            Paragraph paragraph = new Paragraph("\n");
            doc.add(paragraph);
        }
        doc.close();
        writeResponseParameter(response, fileName);
        InputStream in = new ByteArrayInputStream(outputStream.toByteArray());
        WordToPdfUtil.doc2pdf(in, response.getOutputStream());
    }

    @Override
    public void testImport() throws Exception {
        String filePath = "D:\\test.xlsx";
        ImportParams params = new ImportParams();
        params.setHeadRows(2);
        params.setReadSingleCell(true);
        ExcelImportResult<Map> list = ExcelImportUtil.importExcelMore(new File(filePath), Map.class, params);
        Map<String, Object> map = list.getMap();
        System.out.println(JSON.toJSONString(map));

        InputStream inputStream = new FileInputStream(filePath);
        Workbook workbook = null;
        if (filePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (filePath.endsWith(".xls") || filePath.endsWith(".et")) {
            workbook = new HSSFWorkbook(inputStream);
        }
        inputStream.close();
        /**
         * 使用poi读取excel内容
         */
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rows = sheet.rowIterator();
        Row row;
        Cell cell;
        boolean rowStart = false;
        boolean rowStart1= false;
        boolean rowStart2 = false;
        int jumpNum = 0;
        while (rows.hasNext()){
            row = rows.next();
            if (row.getRowNum() < 2){
                String rowValue = getRealType(row.getCell(1));
                System.out.println(rowValue);
                continue;
            }
            Cell cell1 = row.getCell(0);
            String value = getRealType(cell1);
            if ("学生1".equals(value)){
                rowStart = true;
                rowStart1= false;
                rowStart2 = false;
                jumpNum = 2;
            } else if ("学生2".equals(value)){
                rowStart = false;
                rowStart1= true;
                rowStart2 = false;
                jumpNum = 2;
            } else if ("箱货信息".equals(value)){
                rowStart = false;
                rowStart1= false;
                rowStart2 = true;
                jumpNum = 2;
            }
            if (jumpNum != 0 && jumpNum > 0){
                jumpNum--;
                continue;
            }
            // 获取每行的单元格
            Iterator<Cell> cells = row.cellIterator();
            int num = 0;
            while (cells.hasNext()){
                if (rowStart || rowStart1 || rowStart2){
                    cell = cells.next();
                    String cellValue = getRealType(cell);
                    System.out.println(cellValue);
                }
            }
        }
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

    private static String getRealType(Cell cell) {
        if (cell == null) {
            return "";
        }
        String name = cell.getCellType().name();
        if (CellType.STRING.name().contains(name)) {
            return cell.getStringCellValue().toUpperCase();
        }
        if (CellType.NUMERIC.name().contains(name)) {
            return String.valueOf(cell.getNumericCellValue());
        }
        return "";
    }
}
