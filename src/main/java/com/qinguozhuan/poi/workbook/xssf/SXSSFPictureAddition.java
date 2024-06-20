package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * <p>SXSSFPictureAddition is an interface
 * that adds a pictures or a lot of pictures to a {@link SXSSFWorkbook} by once.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 00:09
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SXSSFPictureAddition extends SinglePictureAddition, MultiPictureAddition {


    /**
     * 由于未知原因，放置的对象总是在横向坐标轴有一定的偏移量，此方法针对这个偏移量进行设置，如果不需要偏移量，原值返回即可。
     *
     * @param dx dx1 or dx2
     * @return the offset value, the unit of this data is EMUs
     */
    int increaseTheHorizontalAxisOffset(int dx);

    /**
     * @param workerBookParameter workerBook
     * @param pictureParameter    picture
     * @param sheetParameter      sheet
     * @param coordinate          custom coordinate
     * @param anchor              custom anchor
     */
    void addPicture(WorkerBookParameter workerBookParameter,
                    PictureParameter pictureParameter,
                    SheetParameter sheetParameter,
                    SXSSFCoordinate coordinate,
                    CustomAnchor anchor);

    /**
     * @param sheetParameter sheet
     * @param pictureIndex   the index to this picture (1 based).
     *                       can see method {@link Workbook#addPicture} return.
     * @param coordinate     custom coordinate
     * @param anchor         custom anchor
     */
    void addClientAnchor(SheetParameter sheetParameter,
                         int pictureIndex,
                         SXSSFCoordinate coordinate,
                         CustomAnchor anchor);

}
