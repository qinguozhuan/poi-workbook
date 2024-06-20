package com.qinguozhuan.poi.workbook.xssf;

import lombok.Getter;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/19/2024 00:13
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public final class SinglePictureAdditionParameter implements SXSSFSinglePictureAdditionTemplate {
    @Getter
    private final int maxWidth;//最大宽度
    @Getter
    private final double maxHeight;//最大高度

    @Getter
    private final int topMargins;
    @Getter
    private final int bottomMargins;
    @Getter
    private final int leftMargins;
    @Getter
    private final int rightMargins;

    /**
     * 适合上下等宽，左右等宽模式
     *
     * @param maxWidth    最大宽度 单位 {@link  SXSSFSheet#getColumnWidth(int)}
     * @param maxHeight   最大高度 单位{@link  SXSSFRow#getHeightInPoints()}
     * @param topMargins  上边距，单位pixels
     * @param leftMargins 左边距，单位pixels
     */
    public SinglePictureAdditionParameter(int maxWidth, double maxHeight, int topMargins, int leftMargins) {
        this.maxWidth = maxWidth;
        this.maxHeight = maxHeight;
        this.topMargins = topMargins;
        this.bottomMargins = topMargins;
        this.leftMargins = leftMargins;
        this.rightMargins = leftMargins;
    }

    /**
     * 适合自定义四周边距的模式
     *
     * @param maxWidth      最大宽度 单位 {@link  SXSSFSheet#getColumnWidth(int)}
     * @param maxHeight     最大高度 单位{@link  SXSSFRow#getHeightInPoints()}
     * @param topMargins    上边距，单位pixels
     * @param bottomMargins 下边距，单位pixels
     * @param leftMargins   左边距，单位pixels
     * @param rightMargins  右边距，单位pixels
     */
    public SinglePictureAdditionParameter(int maxWidth,
                                          double maxHeight,
                                          int topMargins,
                                          int bottomMargins,
                                          int leftMargins,
                                          int rightMargins) {
        this.maxWidth = maxWidth;
        this.maxHeight = maxHeight;
        this.topMargins = topMargins;
        this.bottomMargins = bottomMargins;
        this.leftMargins = leftMargins;
        this.rightMargins = rightMargins;
    }
}
