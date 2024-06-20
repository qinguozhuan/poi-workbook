package com.qinguozhuan.poi.workbook.core;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 00:33
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface CustomAnchor extends Anchor {

    /**
     * @return col1 the column (0 based) of the first cell.
     */
    int getCol1();

    /**
     * @return row1 the row (0 based) of the first cell.
     */
    int getRow1();

    /**
     * @return col2 the column (0 based) of the second cell.
     */
    int getCol2();

    /**
     * @return row2 the row (0 based) of the second cell.
     */
    int getRow2();

}
