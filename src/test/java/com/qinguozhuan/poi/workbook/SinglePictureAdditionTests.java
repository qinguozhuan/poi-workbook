package com.qinguozhuan.poi.workbook;

import com.qinguozhuan.poi.workbook.xssf.*;
import lombok.Cleanup;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;


public class SinglePictureAdditionTests {

    private byte[] readFileToByteArray(String filePath) {
        try (InputStream fileInputStream = this.getClass().getResourceAsStream(filePath); ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = fileInputStream.read(buffer)) != -1) {
                byteArrayOutputStream.write(buffer, 0, bytesRead);
            }
            return byteArrayOutputStream.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    @Test
    public void test_insert_a_picture_and_horizontal_alignment_center_in_a_cell_by_bytes() throws IOException {
        int row = 5;
        int col = 5;
        int colWidth = 28 * 256;
        int rowHeight = 50 * 20;

        String url = "/static/width_eq_height(jpeg).png";
        byte[] bytes = this.readFileToByteArray(url);

        SXSSFWorkbook workbook = new SXSSFWorkbook(10);
        SXSSFSheet sheet = workbook.createSheet("test sheet");
        sheet.setColumnWidth(col, colWidth);
        SXSSFRow sheetRow = sheet.createRow(row);
        sheetRow.setHeight((short) rowHeight);

        SXSSFPictureParameter pictureParameter = SXSSFPictureParameter.byBytes(bytes);
        PictureAdditionHelper additionHelper = PictureAdditionHelper.getInstance();
        SinglePictureAdditionParameter singlePictureAddition = new SinglePictureAdditionParameter(sheet.getColumnWidth(col), sheetRow.getHeightInPoints(), 5, 20);
        AnchorParameter anchor = new AnchorParameter(col, row);
        SXSSFCoordinate coordinate = CoordinateParameter.of(singlePictureAddition);
        coordinate.setCoordinate(pictureParameter);
        additionHelper.addPicture(SXSSFWorkerBookParameter.of(workbook),
                pictureParameter,
                SXSSFSheetParameter.of(sheet),
                coordinate,
                anchor);

        //custom
        sheetRow = sheet.createRow(row + 1);
        sheetRow.setHeight((short) rowHeight);
        singlePictureAddition = new SinglePictureAdditionParameter(sheet.getColumnWidth(col), sheetRow.getHeightInPoints(), 5, 20, 11, 10);
        anchor = new AnchorParameter(col, row + 1);
        coordinate = CoordinateParameter.of(singlePictureAddition);
        coordinate.setCoordinate(pictureParameter);
        additionHelper.addPicture(SXSSFWorkerBookParameter.of(workbook),
                pictureParameter,
                SXSSFSheetParameter.of(sheet),
                coordinate,
                anchor);

        String userHome = System.getProperty("user.home");
        @Cleanup FileOutputStream out = new FileOutputStream(userHome + "/Downloads/test_excel_" + System.currentTimeMillis() + ".xlsx");
        workbook.write(out);
        workbook.close();
    }

    @Test
    public void test_insert_a_picture_and_horizontal_alignment_center_in_a_cell_by_url() throws IOException {
        int row = 5;
        int col = 5;
        int colWidth = 28 * 256;
        int rowHeight = 50 * 20;

        String url = "https://bkimg.cdn.bcebos.com/pic/8694a4c27d1ed21b0ef480aab225cac451da81cb9163?x-bce-process=image";

        SXSSFWorkbook workbook = new SXSSFWorkbook(10);
        SXSSFSheet sheet = workbook.createSheet("test sheet");
        sheet.setColumnWidth(col, colWidth);
        SXSSFRow sheetRow = sheet.createRow(row);
        sheetRow.setHeight((short) rowHeight);

        SXSSFPictureParameter pictureParameter = SXSSFPictureParameter.byURL(url);
        PictureAdditionHelper additionHelper = PictureAdditionHelper.getInstance();
        SinglePictureAdditionParameter singlePictureAddition = new SinglePictureAdditionParameter(sheet.getColumnWidth(col), sheetRow.getHeightInPoints(), 5, 20);
        AnchorParameter anchor = new AnchorParameter(col, row);
        SXSSFCoordinate coordinate = CoordinateParameter.of(singlePictureAddition);
        coordinate.setCoordinate(pictureParameter);
        additionHelper.addPicture(SXSSFWorkerBookParameter.of(workbook),
                pictureParameter,
                SXSSFSheetParameter.of(sheet),
                coordinate,
                anchor);

        //custom
        sheetRow = sheet.createRow(row + 1);
        sheetRow.setHeight((short) rowHeight);
        singlePictureAddition = new SinglePictureAdditionParameter(sheet.getColumnWidth(col), sheetRow.getHeightInPoints(), 5, 20, 11, 10);
        anchor = new AnchorParameter(col, row + 1);
        coordinate = CoordinateParameter.of(singlePictureAddition);
        coordinate.setCoordinate(pictureParameter);
        additionHelper.addPicture(SXSSFWorkerBookParameter.of(workbook),
                pictureParameter,
                SXSSFSheetParameter.of(sheet),
                coordinate,
                anchor);

        String userHome = System.getProperty("user.home");
        @Cleanup FileOutputStream out = new FileOutputStream(userHome + "/Downloads/test_excel_" + System.currentTimeMillis() + ".xlsx");
        workbook.write(out);
        workbook.close();
    }

}
