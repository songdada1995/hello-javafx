package org.openjfx.model;

import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author songbo
 * @version 1.0
 * @date 2023/10/22 12:07
 */
@ExcelIgnoreUnannotated
@Data
public class JobTask1ExcelModelRead {

    @ExcelProperty("Invoice ID")
    private String invoiceID;

    @ExcelProperty("Original balance")
    private String originalBalance;

}
