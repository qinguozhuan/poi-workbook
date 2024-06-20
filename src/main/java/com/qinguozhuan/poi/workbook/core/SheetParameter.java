package com.qinguozhuan.poi.workbook.core;

import org.apache.poi.ss.usermodel.Sheet;

import java.io.Serializable;

/**
 * <p>Packaging for {@link Sheet} parameters.
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/18/2024 00:19
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
public interface SheetParameter extends Serializable {

    /**
     * Get sheet
     *
     * @return can see {@link Sheet}.
     */
    Sheet getSheet();

}
