package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.PictureAdditionTemplate;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * <p>Template for adding pictures.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 23:48
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SXSSFPictureAdditionTemplate extends PictureAdditionTemplate {

    /**
     * Get the max width of the container being added.
     *
     * @return The unit of this data is {@link  Sheet#getColumnWidth(int)}.
     */
    int getMaxWidth();

    /**
     * Get the max height of the container being added.
     *
     * @return The unit of this data is {@link  Row#getHeightInPoints()}.
     */
    double getMaxHeight();

}
