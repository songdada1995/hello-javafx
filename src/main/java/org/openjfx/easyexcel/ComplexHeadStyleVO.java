package org.openjfx.easyexcel;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.openjfx.easyexcel.style.HeadWriteCellStyle;

/**
 * 复杂表头样式信息，包含需要自定义的表头坐标及样式
 *
 * @author hukai
 * @date 2022/3/16
 */
@Data
@AllArgsConstructor
public class ComplexHeadStyleVO {

    /**
     * 表头横坐标 - 行
     */
    private Integer x;

    /**
     * 表头纵坐标 - 列
     */
    private Integer y;

    /**
     * 属性
     */
    private HeadWriteCellStyle headWriteCellStyle;
}