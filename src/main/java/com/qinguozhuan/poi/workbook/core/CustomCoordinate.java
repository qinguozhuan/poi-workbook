package com.qinguozhuan.poi.workbook.core;

import org.apache.poi.ss.usermodel.ChildAnchor;

/**
 * <p>Common interface for anchors.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 00:24
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface CustomCoordinate extends Anchor {

    /**
     * can see {@link ChildAnchor#getDx1()}
     *
     * @return The unit of this data is EMU
     */
    int getDx1();

    /**
     * can see {@link ChildAnchor#getDy1()}
     *
     * @return The unit of this data is EMU
     */
    int getDy1();

    /**
     * can see {@link ChildAnchor#getDx2()}
     *
     * @return The unit of this data is EMU
     */
    int getDx2();

    /**
     * can see {@link ChildAnchor#getDy2()}
     *
     * @return The unit of this data is EMU
     */
    int getDy2();
}
