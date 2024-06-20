package com.qinguozhuan.poi.workbook.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * <p>MultiPictureAddition is an interface
 * that adds a lot of pictures to a {@link Workbook} by once.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/17/2024 23:10
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface MultiPictureAddition extends PictureAddition {

    /**
     * Adds multi picture to a workbook. one by one insert.
     *
     * @param workerBookParameter Contains customized parameters for class {@link Workbook}.
     * @param pictureParameters   {@link PictureParameter} list.
     * @return the index to this picture (1 based). sort order by {@param pictureFiles} list index order.
     */
    List<Integer> addPicture(WorkerBookParameter workerBookParameter, List<PictureParameter> pictureParameters);

}
