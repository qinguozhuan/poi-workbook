package com.qinguozhuan.poi.workbook.xssf;

/**
 * <p>Template for adding multi pictures.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 23:51
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SXSSFMultiFixedSizePictureAdditionTemplate extends SXSSFPictureAdditionTemplate {

    /**
     * Get the total pictures being added.
     *
     * @return total number
     */
    int getTotal();

    /**
     * Get the fixed width of the picture being added.
     *
     * @return The unit of this data is inch.
     */
    double getFixedWidth();

    /**
     * Get the fixed height of the picture being added.
     *
     * @return The unit of this data is inch.
     */
    double getFixedHeight();


    /**
     * Get each picture spacing.
     *
     * @return The unit of this data is EMUs.
     */
    int getPictureSpacing();

}
