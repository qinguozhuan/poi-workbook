package com.qinguozhuan.poi.workbook.core;


import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p>SinglePictureAddition is an interface
 * that adds a single picture to a {@link Workbook} by once.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/17/2024 23:10
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SinglePictureAddition extends PictureAddition {

    /**
     * Adds a picture to the workbook.Extended for
     * {@link Workbook#addPicture(byte[], int)} methods
     *
     * @param workerBookParameter Contains customized parameters for class {@link Workbook}.
     * @param pictureParameter    {@link PictureParameter}.
     * @return the index to this picture (1 based).
     */
    int addPicture(WorkerBookParameter workerBookParameter, PictureParameter pictureParameter);

}
