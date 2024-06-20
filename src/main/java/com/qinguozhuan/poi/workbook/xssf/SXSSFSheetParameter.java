package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.SheetParameter;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.Serial;
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
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public final class SXSSFSheetParameter implements SheetParameter {
    /**
     * use serialVersionUID from JDK 1.0.2 for interoperability
     */
    @Serial
    private static final long serialVersionUID = -8742448824652078965L;

    @Getter
    private final SXSSFSheet sheet;

    /**
     * build object by SXSSFSheet
     *
     * @param sheet SXSSFSheet
     * @return SXSSFSheetParameter
     */
    public static SXSSFSheetParameter of(SXSSFSheet sheet) {
        Objects.requireNonNull(sheet, "sheet must not be null.");
        return new SXSSFSheetParameter(sheet);
    }

}
