package com.qinguozhuan.poi.workbook;

import com.qinguozhuan.poi.workbook.xssf.*;
import lombok.Cleanup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/20/2024 22:38
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public class MultiPictureAdditionTests {

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
    public void test_insert_multi_pictures_into_a_cell_in_the_center_by_bytes() throws IOException {
        int row = 5;
        int col = 5;
        int colWidth = 29 * 255;
        int rowHeight = 170 * 20;
        int total = 5;

        String url = "/static/test_tag.jpeg";
        byte[] bytes = this.readFileToByteArray(url);

        SXSSFWorkbook workbook = new SXSSFWorkbook(10);
        SXSSFSheet sheet = workbook.createSheet("test sheet");
        sheet.setColumnWidth(col, colWidth);
        Row sheetRow = sheet.createRow(row);
        sheetRow.setHeight((short) rowHeight);
        sheetRow = sheet.createRow(row + 1);
        sheetRow.setHeight((short) rowHeight);
        sheetRow = sheet.createRow(row + 2);
        sheetRow.setHeight((short) rowHeight);

        SXSSFPictureParameter pictureParameter = SXSSFPictureParameter.byBytes(bytes);
        SXSSFWorkerBookParameter workerBookParameter = SXSSFWorkerBookParameter.of(workbook);
        SXSSFSheetParameter sheetParameter = SXSSFSheetParameter.of(sheet);
        PictureAdditionHelper pictureAdditionHelper = PictureAdditionHelper.getInstance();

        MultiFixedSizePictureAdditionParameter multiPictureAddition = new MultiFixedSizePictureAdditionParameter(sheet.getColumnWidth(col), sheetRow.getHeightInPoints(), total, 0.23, 0.21);
        AnchorParameter anchor = new AnchorParameter(col, row);
        for (int i = 0; i <= total; i++) {
            SXSSFCoordinate coordinate = CoordinateParameter.of(multiPictureAddition);
            coordinate.setCoordinate(i);

            pictureAdditionHelper.addPicture(workerBookParameter, pictureParameter, sheetParameter, coordinate, anchor);
        }

        String userHome = System.getProperty("user.home");
        @Cleanup FileOutputStream out = new FileOutputStream(userHome + "/Downloads/test_excel_" + System.currentTimeMillis() + ".xlsx");
        workbook.write(out);
        workbook.close();
    }
}
