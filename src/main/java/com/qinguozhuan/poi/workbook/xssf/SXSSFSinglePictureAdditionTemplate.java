package com.qinguozhuan.poi.workbook.xssf;

/**
 * <p>Template for adding a pictures.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 23:50
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SXSSFSinglePictureAdditionTemplate extends SXSSFPictureAdditionTemplate {

    /**
     * Get top margins
     *
     * @return The unit of this data is pixel.
     */
    int getTopMargins();

    /**
     * Get bottom margins
     *
     * @return The unit of this data is pixel.
     */
    int getBottomMargins();

    /**
     * Get left margins
     *
     * @return The unit of this data is pixel.
     */
    int getLeftMargins();

    /**
     * Get right margins
     *
     * @return The unit of this data is pixel.
     */
    int getRightMargins();

}
