package org.openjfx.easyexcel;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.openjfx.easyexcel.style.ContentWriteCellStyle;

/**
 * 自定义内容样式
 *
 * @author songbo
 * @date 2023-05-29
 */
@Data
@AllArgsConstructor
public class ComplexContentStyleVO {

    /**
     * 横坐标 - 行
     */
    private Integer x;

    /**
     * 纵坐标 - 列
     */
    private Integer y;

    /**
     * 属性
     */
    private ContentWriteCellStyle contentWriteCellStyle;
}