package com.qinguozhuan.poi.workbook.core;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Serializable;

/**
 * <p>Packaging for {@link Workbook} parameters.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/17/2024 23:21
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface WorkerBookParameter extends Serializable {

    /**
     * Get sheet
     *
     * @return can see {@link Sheet}.
     */
    Workbook getWorkbook();

    /**
     * @param index
     * @return SheetParameter
     * @see Workbook#getSheetAt(int)
     */
    SheetParameter getSheetParameter(int index);

}
