package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.PictureParameter;
import com.qinguozhuan.poi.workbook.utils.ImageUtil;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.ChildAnchor;
import org.apache.poi.util.Units;
import org.springframework.util.Assert;

import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.Serial;
import java.io.Serializable;
import java.util.Objects;

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
@Slf4j
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public final class CoordinateParameter implements SXSSFCoordinate, Serializable {
    /**
     * use serialVersionUID from JDK 1.0.2 for interoperability
     */
    @Serial
    private static final long serialVersionUID = -8742448824652078965L;

    @Getter
    private int dx1;
    @Getter
    private int dx2;
    @Getter
    private int dy1;
    @Getter
    private int dy2;

    private final SXSSFSinglePictureAdditionTemplate singlePictureAdditionTemplate;
    private final SXSSFMultiFixedSizePictureAdditionTemplate multiFixedSizePictureAdditionTemplate;

    private CoordinateParameter(SXSSFSinglePictureAdditionTemplate singlePictureAdditionTemplate,
                                SXSSFMultiFixedSizePictureAdditionTemplate multiFixedSizePictureAdditionTemplate) {
        this.singlePictureAdditionTemplate = singlePictureAdditionTemplate;
        this.multiFixedSizePictureAdditionTemplate = multiFixedSizePictureAdditionTemplate;
    }

    /**
     * 合适放置上下等距，左右等距的单张图片
     *
     * @param singlePictureAdditionTemplate 模式
     */
    public static CoordinateParameter of(SXSSFSinglePictureAdditionTemplate singlePictureAdditionTemplate) {
        Objects.requireNonNull(singlePictureAdditionTemplate,
                "singlePictureAdditionTemplate must not be null.");

        return new CoordinateParameter(singlePictureAdditionTemplate, null);
    }

    /**
     * 适合放置多个固定图标，且剧中放置，图标间距平均的设置
     *
     * @param multiFixedSizePictureAdditionTemplate 模式
     */
    public static CoordinateParameter of(SXSSFMultiFixedSizePictureAdditionTemplate multiFixedSizePictureAdditionTemplate) {
        Objects.requireNonNull(multiFixedSizePictureAdditionTemplate,
                "multiFixedSizePictureAdditionTemplate must not be null.");

        return new CoordinateParameter(null, multiFixedSizePictureAdditionTemplate);
    }

    /**
     * @see ChildAnchor
     */
    public enum AnchorEnum {
        DX1, DY1, DX2, DY2
    }

    @Override
    public SXSSFCoordinate setCoordinate(PictureParameter pictureParameter) throws IOException {
        Objects.requireNonNull(singlePictureAdditionTemplate,
                "Not supported in current mode, singlePictureAdditionTemplate must not be null!");

        byte[] bytes = pictureParameter.getBytes();
        BufferedImage image = ImageUtil.toBufferedImage(bytes);
        calcCoordinate(image);
        return this;
    }

    @Override
    public SXSSFCoordinate setCoordinate(int index) {
        Objects.requireNonNull(multiFixedSizePictureAdditionTemplate,
                "Not supported in current mode, multiFixedSizePictureAdditionTemplate must not be null!");
        calcCoordinate(index);
        return this;
    }

    /**
     * calculate single picture mode coordinate
     *
     * @param image {@link BufferedImage}
     */
    private void calcCoordinate(BufferedImage image) {
        Objects.requireNonNull(singlePictureAdditionTemplate,
                "Not supported in current mode, singlePictureAdditionTemplate must not be null!");
        Objects.requireNonNull(image, "Image must not be null!");

        final int TOP_MARGINS = Units.pixelToEMU(singlePictureAdditionTemplate.getTopMargins());//EMU
        final int BOTTOM_MARGINS = Units.pixelToEMU(singlePictureAdditionTemplate.getBottomMargins());//EMU
        final int LEFT_MARGINS = Units.pixelToEMU(singlePictureAdditionTemplate.getLeftMargins());//EMU
        final int RIGHT_MARGINS = Units.pixelToEMU(singlePictureAdditionTemplate.getRightMargins());//EMU
        final int MAX_WIDTH = Units.columnWidthToEMU(singlePictureAdditionTemplate.getMaxWidth());//EMU
        final int MAX_HEIGHT = Units.toEMU(singlePictureAdditionTemplate.getMaxHeight());//EMU

        final double CONTAINER_WIDTH = MAX_WIDTH - LEFT_MARGINS - RIGHT_MARGINS;//EMU
        final double CONTAINER_HEIGHT = MAX_HEIGHT - TOP_MARGINS - BOTTOM_MARGINS;//EMU
        final double IMAGE_WIDTH = Units.pixelToEMU(image.getWidth()) * 1.0D;//EMU
        final double IMAGE_HEIGHT = Units.pixelToEMU(image.getHeight()) * 1.0D;//EMU

        //Calculate the aspect ratio of A and B
        final double IMAGE_RATIO = IMAGE_WIDTH / IMAGE_HEIGHT;
        final double CONTAINER_RATIO = CONTAINER_WIDTH / CONTAINER_HEIGHT;

        log.info("Setting value = {\nTOP_MARGINS={};" +
                        "\nBOTTOM_MARGINS={}" +
                        "\nLEFT_MARGINS={}" +
                        "\nRIGHT_MARGINS={}" +
                        "\nMAX_WIDTH={}" +
                        "\nMAX_HEIGHT={}" +
                        "\nCONTAINER_WIDTH={}" +
                        "\nCONTAINER_HEIGHT={}" +
                        "\nIMAGE_WIDTH={}" +
                        "\nIMAGE_HEIGHT={}\n}",
                TOP_MARGINS,
                BOTTOM_MARGINS,
                LEFT_MARGINS,
                RIGHT_MARGINS,
                MAX_WIDTH,
                MAX_HEIGHT,
                CONTAINER_WIDTH,
                CONTAINER_HEIGHT,
                IMAGE_WIDTH,
                IMAGE_HEIGHT);

        //Compare the aspect ratios of A and B to determine how to place them
        double actImageWidth, actImageHeight;
        if (IMAGE_RATIO > CONTAINER_RATIO) {
            //Scale image's dimensions to fit container's width
            actImageWidth = CONTAINER_WIDTH;
            actImageHeight = (CONTAINER_WIDTH / IMAGE_WIDTH) * IMAGE_HEIGHT;

            //Check that the scaled height does not exceed the height of container
            if (actImageHeight > CONTAINER_HEIGHT) {
                actImageHeight = CONTAINER_HEIGHT;
                actImageWidth = (CONTAINER_HEIGHT / IMAGE_HEIGHT) * IMAGE_WIDTH;
            }
        } else if (IMAGE_RATIO < CONTAINER_RATIO) {
            //Scale image's dimensions to fit container's height
            actImageWidth = (CONTAINER_HEIGHT / IMAGE_HEIGHT) * IMAGE_WIDTH;
            actImageHeight = CONTAINER_HEIGHT;

            //Check that the scaled width exceeds the width of container
            if (actImageWidth > CONTAINER_WIDTH) {
                actImageWidth = CONTAINER_WIDTH;
                actImageHeight = (CONTAINER_WIDTH / IMAGE_WIDTH) * IMAGE_HEIGHT;
            }
        } else {
            actImageWidth = CONTAINER_WIDTH;
            actImageHeight = CONTAINER_HEIGHT;
        }
        double actTopMargins, actBottomMargins;
        if (Math.abs(TOP_MARGINS) + Math.abs(BOTTOM_MARGINS) + actImageHeight < MAX_HEIGHT) {
            actTopMargins = ((MAX_HEIGHT - actImageHeight) / (TOP_MARGINS + BOTTOM_MARGINS)) * TOP_MARGINS;
            actBottomMargins = ((MAX_HEIGHT - actImageHeight) / (TOP_MARGINS + BOTTOM_MARGINS)) * BOTTOM_MARGINS;
        } else {
            actTopMargins = TOP_MARGINS;
            actBottomMargins = BOTTOM_MARGINS;
        }
        double actLeftMargins, actRightMargins;
        if (Math.abs(LEFT_MARGINS) + Math.abs(RIGHT_MARGINS) + actImageWidth < MAX_WIDTH) {
            actLeftMargins = ((MAX_WIDTH - actImageWidth) / (LEFT_MARGINS + RIGHT_MARGINS)) * LEFT_MARGINS;
            actRightMargins = ((MAX_WIDTH - actImageWidth) / (LEFT_MARGINS + RIGHT_MARGINS)) * RIGHT_MARGINS;
        } else {
            actLeftMargins = LEFT_MARGINS;
            actRightMargins = RIGHT_MARGINS;
        }

        log.info("Actual value = {\nactImageWidth={};" +
                        "\nactImageHeight={}" +
                        "\nactTopMargins={}" +
                        "\nactBottomMargins={}" +
                        "\nactLeftMargins={}" +
                        "\nactRightMargins={}\n}",
                actImageWidth,
                actImageHeight,
                actTopMargins,
                actBottomMargins,
                actLeftMargins,
                actRightMargins);

        for (AnchorEnum anchor : AnchorEnum.values()) {
            switch (anchor) {
                case DX1 -> this.dx1 = (int) Math.ceil(actLeftMargins);
                case DY1 -> this.dy1 = (int) Math.ceil(actTopMargins);
                case DX2 -> this.dx2 = (int) Math.ceil(actLeftMargins)
                        + (int) Math.ceil(actImageWidth);
                case DY2 -> this.dy2 = (int) Math.ceil(actTopMargins)
                        + (int) Math.ceil(actImageHeight);
            }
        }
    }

    /**
     * calculate multi fixed picture mode Coordinate
     *
     * @param index the index of pictures, must more than -1;
     */
    private void calcCoordinate(int index) {
        Assert.isTrue(index >= 0, "index must be greater than or equal to zero.");
        Objects.requireNonNull(multiFixedSizePictureAdditionTemplate,
                "Not supported in current mode, multiFixedSizePictureAdditionTemplate must not be null!");

        //Fixed picture size
        final int PICTURE_WIDTH = Units.toEMU(multiFixedSizePictureAdditionTemplate.getFixedWidth() * Units.POINT_DPI);//EMU
        final int PICTURE_HEIGHT = Units.toEMU(multiFixedSizePictureAdditionTemplate.getFixedHeight() * Units.POINT_DPI);//EMU
        final int CONTAINER_WIDTH = Units.columnWidthToEMU(multiFixedSizePictureAdditionTemplate.getMaxWidth());//EMU
        final int CONTAINER_HEIGHT = Units.toEMU(multiFixedSizePictureAdditionTemplate.getMaxHeight());//EMU

        //each picture spacing
        final int PICTURE_SPACING = multiFixedSizePictureAdditionTemplate.getPictureSpacing();
        final int TOTAL = multiFixedSizePictureAdditionTemplate.getTotal();
        //container width = margin * 2 + number of pictures * picture width + each picture spacing * (total number of pictures - 1);
        final double WIDTH_MARGINS = (CONTAINER_WIDTH - TOTAL * PICTURE_WIDTH - PICTURE_SPACING * (TOTAL - 1)) / 2D;
        //container height = margin * 2 + pictures height;
        final double HEIGHT_MARGINS = (CONTAINER_HEIGHT - PICTURE_HEIGHT) / 2D;

        log.info("Setting value = {\nPICTURE_WIDTH={};" +
                        "\nPICTURE_HEIGHT={}" +
                        "\nCONTAINER_WIDTH={}" +
                        "\nCONTAINER_HEIGHT={}" +
                        "\nPICTURE_SPACING={}" +
                        "\nTOTAL={}" +
                        "\nWIDTH_MARGINS={}" +
                        "\nHEIGHT_MARGINS={}\n}",
                PICTURE_WIDTH,
                PICTURE_HEIGHT,
                CONTAINER_WIDTH,
                CONTAINER_HEIGHT,
                PICTURE_SPACING,
                TOTAL,
                WIDTH_MARGINS,
                HEIGHT_MARGINS);

        for (AnchorEnum anchor : AnchorEnum.values()) {
            switch (anchor) {
                case DX1 -> this.dx1 = (int) Math.ceil(WIDTH_MARGINS + index * (PICTURE_WIDTH + PICTURE_SPACING));
                case DY1 -> this.dy1 = (int) Math.ceil(HEIGHT_MARGINS);
                case DX2 -> this.dx2 = (int) Math.ceil(PICTURE_WIDTH
                        + (WIDTH_MARGINS + index * (PICTURE_WIDTH + PICTURE_SPACING)));
                case DY2 -> this.dy2 = (int) Math.ceil(PICTURE_HEIGHT + HEIGHT_MARGINS);
            }
        }
    }

}
