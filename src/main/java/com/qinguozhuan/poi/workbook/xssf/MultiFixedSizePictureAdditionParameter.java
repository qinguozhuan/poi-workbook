package com.qinguozhuan.poi.workbook.xssf;

import lombok.Getter;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.springframework.util.Assert;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/19/2024 00:05
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public final class MultiFixedSizePictureAdditionParameter implements SXSSFMultiFixedSizePictureAdditionTemplate {
    @Getter
    private final int maxWidth;//最大宽度
    @Getter
    private final double maxHeight;//最大高度

    @Getter
    private final int total;//总数
    @Getter
    private final double fixedWidth;//英寸
    @Getter
    private final double fixedHeight; //英寸
    @Getter
    private final int pictureSpacing;//EMU

    /**
     * @param maxWidth    最大宽度 单位 {@link  SXSSFSheet#getColumnWidth(int)}
     * @param maxHeight   最大高度 单位{@link  SXSSFRow#getHeightInPoints()}
     * @param total       总数
     * @param fixedWidth  固定宽度 单位英寸
     * @param fixedHeight 固定高度 单位英寸
     */
    public MultiFixedSizePictureAdditionParameter(int maxWidth,
                                                  double maxHeight,
                                                  int total,
                                                  double fixedWidth,
                                                  double fixedHeight) {
        this.maxWidth = maxWidth;
        this.maxHeight = maxHeight;
        this.total = total;
        this.fixedWidth = fixedWidth;
        this.fixedHeight = fixedHeight;
        this.pictureSpacing = Units.toEMU(fixedWidth * Units.POINT_DPI) / 4;//EMU
    }

    /**
     * @param maxWidth       最大宽度 单位 {@link  SXSSFSheet#getColumnWidth(int)}
     * @param maxHeight      最大高度 单位{@link  SXSSFRow#getHeightInPoints()}
     * @param total          总数
     * @param fixedWidth     固定宽度 单位英寸
     * @param fixedHeight    固定高度 单位英寸
     * @param pictureSpacing 图片间距，单位pixel
     */
    public MultiFixedSizePictureAdditionParameter(int maxWidth,
                                                  double maxHeight,
                                                  int total,
                                                  double fixedWidth,
                                                  double fixedHeight,
                                                  int pictureSpacing) {
        Assert.isTrue(maxWidth > 0, "maxWidth must be greater than zero.");
        Assert.isTrue(maxHeight > 0, "maxHeight must be greater than zero.");
        Assert.isTrue(total > 0, "total must be greater than zero.");
        Assert.isTrue(fixedWidth > 0, "fixedWidth must be greater than zero.");
        Assert.isTrue(fixedHeight > 0, "fixedWidth must be greater than zero.");
        Assert.isTrue(pictureSpacing > 0, "pictureSpacing must be greater than zero.");

        this.maxWidth = maxWidth;
        this.maxHeight = maxHeight;
        this.total = total;
        this.fixedWidth = fixedWidth;
        this.fixedHeight = fixedHeight;
        this.pictureSpacing = Units.pixelToEMU(pictureSpacing);
    }

}
