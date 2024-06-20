package com.qinguozhuan.poi.workbook.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.Serializable;

/**
 * <p>Packaging for image parameters.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/17/2024 23:23
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface PictureParameter extends Serializable {

    /**
     * Get picture data
     *
     * @return The bytes data of a picture
     */
    byte[] getBytes();

    /**
     * Get picture type of Workbook type
     *
     * @return Follow as values
     * <ol>
     * <li>{@link Workbook#PICTURE_TYPE_EMF }
     * <li>{@link Workbook#PICTURE_TYPE_WMF}
     * <li>{@link Workbook#PICTURE_TYPE_PICT}
     * <li>{@link Workbook#PICTURE_TYPE_JPEG}
     * <li>{@link Workbook#PICTURE_TYPE_PNG}
     * <li>{@link Workbook#PICTURE_TYPE_DIB}
     * </ol>
     */
    int getFormat();

}
