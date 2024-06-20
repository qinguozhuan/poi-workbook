package com.qinguozhuan.poi.workbook.xssf;

import com.qinguozhuan.poi.workbook.core.SheetParameter;
import com.qinguozhuan.poi.workbook.core.WorkerBookParameter;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.Assert;

import java.io.Serial;
import java.util.Objects;

/**
 * <p>
 *
 * <p>If you have any questions, let me know!
 * My email is qgzh_123@163.com.
 *
 * <p>Creation time: 06/17/2024 23:43
 *
 * @author Jon.Qin
 * @version 1.0.0
 */
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public final class SXSSFWorkerBookParameter implements WorkerBookParameter {
    /**
     * use serialVersionUID from JDK 1.0.2 for interoperability
     */
    @Serial
    private static final long serialVersionUID = -8742448824652078965L;

    @Getter
    private final SXSSFWorkbook workbook;

    @Override
    public SheetParameter getSheetParameter(int index) {
        Assert.isTrue(index >= 0, "Index must be greater than or equal to zero.");
        return SXSSFSheetParameter.of(workbook.getSheetAt(index));
    }

    /**
     * build object by SXSSFWorkbook
     *
     * @param workbook SXSSFWorkbook
     * @return SXSSFWorkerBookParameter
     */
    public static SXSSFWorkerBookParameter of(SXSSFWorkbook workbook) {
        Objects.requireNonNull(workbook, "workbook must not be null.");
        return new SXSSFWorkerBookParameter(workbook);
    }

}
