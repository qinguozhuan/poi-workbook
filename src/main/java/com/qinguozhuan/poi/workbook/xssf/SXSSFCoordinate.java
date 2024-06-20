package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.CustomCoordinate;
import com.qinguozhuan.poi.workbook.core.PictureParameter;

import java.io.IOException;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/19/2024 21:57
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SXSSFCoordinate extends CustomCoordinate {

    /**
     * 设置布局,适用于单张图片模式
     *
     * @param pictureParameter
     * @throws IOException
     */
    SXSSFCoordinate setCoordinate(PictureParameter pictureParameter) throws IOException;

    /**
     * 设置布局,适用于多张固定图片模式
     *
     * @param index 图片下标索引，从0开始
     */
    SXSSFCoordinate setCoordinate(int index);
}
