package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.CustomAnchor;
import lombok.Getter;
import org.springframework.util.Assert;

import java.io.Serial;
import java.io.Serializable;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 00:18
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public final class AnchorParameter implements CustomAnchor, Serializable {

    /**
     * use serialVersionUID from JDK 1.0.2 for interoperability
     */
    @Serial
    private static final long serialVersionUID = -8742448824652078965L;

    @Getter
    private final int col1;
    @Getter
    private final int col2;
    @Getter
    private final int row1;
    @Getter
    private final int row2;

    /**
     * custom location
     *
     * @param col1
     * @param col2
     * @param row1
     * @param row2
     */
    public AnchorParameter(int col1, int col2, int row1, int row2) {
        Assert.isTrue(col1 >= 0, "col1 must be greater than or equal to zero.");
        Assert.isTrue(col2 >= 0, "col2 must be greater than or equal to zero.");
        Assert.isTrue(row1 >= 0, "row1 must be greater than or equal to zero.");
        Assert.isTrue(row2 >= 0, "row2 must be greater than or equal to zero.");

        this.col1 = col1;
        this.col2 = col2;
        this.row1 = row1;
        this.row2 = row2;
    }

    /**
     * in a cell
     *
     * @param col
     * @param row
     */
    public AnchorParameter(int col, int row) {
        this(col, col, row, row);
    }
}
