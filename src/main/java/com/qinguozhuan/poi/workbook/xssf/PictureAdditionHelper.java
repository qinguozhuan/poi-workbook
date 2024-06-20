package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.CustomAnchor;
import com.qinguozhuan.poi.workbook.core.PictureParameter;
import com.qinguozhuan.poi.workbook.core.SheetParameter;
import com.qinguozhuan.poi.workbook.core.WorkerBookParameter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.springframework.util.Assert;

import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/17/2024 23:41
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
@Slf4j
public final class PictureAdditionHelper implements SXSSFPictureAddition {

    private PictureAdditionHelper() {
    }

    private final static class BeanShop {
        private static final PictureAdditionHelper INSTANCE = new PictureAdditionHelper();
    }

    /**
     * Get PictureAdditionHelper instance
     *
     * @return PictureAdditionHelper
     */
    public static PictureAdditionHelper getInstance() {
        return BeanShop.INSTANCE;
    }

    /**
     * The relative position offset coefficient, the left and right margins must be maintained with coefficients
     */
    private static final double OFFSET_CONSTANT = 1.15D;

    @Override
    public int increaseTheHorizontalAxisOffset(int dx) {
        return (int) Math.ceil(dx * OFFSET_CONSTANT);
    }

    @Override
    public List<Integer> addPicture(WorkerBookParameter workerBookParameter, List<PictureParameter> pictureParameters) {
        Objects.requireNonNull(workerBookParameter, "The workerBookParameter cannot be empty.");

        if (pictureParameters == null || pictureParameters.isEmpty()) {
            return Collections.emptyList();
        }

        return pictureParameters.stream()
                .map(pictureParameter -> this.addPicture(workerBookParameter, pictureParameter))
                .collect(Collectors.toList());
    }

    @Override
    public int addPicture(WorkerBookParameter workerBookParameter, PictureParameter pictureParameter) {
        Objects.requireNonNull(workerBookParameter, "The workerBookParameter cannot be empty.");
        Objects.requireNonNull(pictureParameter, "The pictureParameter cannot be empty.");

        return workerBookParameter.getWorkbook().addPicture(pictureParameter.getBytes(), pictureParameter.getFormat());
    }

    @Override
    public void addPicture(WorkerBookParameter workerBookParameter,
                           PictureParameter pictureParameter,
                           SheetParameter sheetParameter,
                           SXSSFCoordinate coordinate,
                           CustomAnchor anchor) {
        Objects.requireNonNull(workerBookParameter, "The workerBookParameter cannot be empty.");
        Objects.requireNonNull(pictureParameter, "The pictureParameter cannot be empty.");
        Objects.requireNonNull(sheetParameter, "The sheetParameter cannot be empty.");
        Objects.requireNonNull(coordinate, "The coordinate cannot be empty.");
        Objects.requireNonNull(anchor, "The anchor cannot be empty.");

        int pictureIndex = this.addPicture(workerBookParameter, pictureParameter);
        this.addClientAnchor(sheetParameter, pictureIndex, coordinate, anchor);
    }

    @Override
    public void addClientAnchor(SheetParameter sheetParameter,
                                int pictureIndex,
                                SXSSFCoordinate coordinate,
                                CustomAnchor anchor) {
        Objects.requireNonNull(sheetParameter, "The sheetParameter cannot be empty.");
        Assert.isTrue(pictureIndex >= 0, "pictureIndex must be greater than or equal to zero.");
        Objects.requireNonNull(coordinate, "The coordinate cannot be empty.");
        Objects.requireNonNull(anchor, "The anchor cannot be empty.");

        int dx1 = coordinate.getDx1();
        int dx2 = coordinate.getDx2();
        int dy1 = coordinate.getDy1();
        int dy2 = coordinate.getDy2();
        log.info("{\ndx1={};\ndx2={};\ndy1={};\ndy2={}\n}", dx1, dx2, dy1, dy2);
        Assert.isTrue(!(dx1 == 0 && dx2 == 0 && dy1 == 0 && dy2 == 0),
                "Please calculate the value of coordinate before add ClientAnchor.");

        int col1 = anchor.getCol1();
        int col2 = anchor.getCol2();
        int row1 = anchor.getRow1();
        int row2 = anchor.getRow2();

        ClientAnchor clientAnchor = new XSSFClientAnchor(
                this.increaseTheHorizontalAxisOffset(dx1),
                dy1,
                this.increaseTheHorizontalAxisOffset(dx2),
                dy2,
                col1,
                row1,
                col2,
                row2);
        clientAnchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
        Drawing<?> drawingPatriarch = sheetParameter.getSheet().createDrawingPatriarch();
        drawingPatriarch.createPicture(clientAnchor, pictureIndex);
    }
}
