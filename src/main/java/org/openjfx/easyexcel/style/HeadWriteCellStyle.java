package org.openjfx.easyexcel.style;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import lombok.Data;
import org.openjfx.constants.EasyExcelConstants;

/**
 * @author songbo
 * @version 1.0
 * @date 2023/5/29 19:14
 */
@Data
public class HeadWriteCellStyle extends WriteCellStyle {

    /**
     * 列宽
     */
    private Integer columnWidth;

    /**
     * 行高
     */
    private Short rowHeight;

    /**
     * 默认头部样式
     *
     * @return
     */
    public static HeadWriteCellStyle getDefaultHeadStyle() {
        HeadWriteCellStyle headStyle = new HeadWriteCellStyle();
        WriteFont writeFont = new WriteFont();
        writeFont.setBold(false);
        writeFont.setFontName(EasyExcelConstants.DefaultConstants.DEFAULT_FONT_NAME);
        writeFont.setFontHeightInPoints(EasyExcelConstants.DefaultConstants.DEFAULT_FONT_SIZE);
        headStyle.setWriteFont(writeFont);
        headStyle.setColumnWidth(EasyExcelConstants.DefaultConstants.DEFAULT_COLUMN_WIDTH);
        headStyle.setRowHeight(EasyExcelConstants.DefaultConstants.DEFAULT_ROW_HEIGHT);
        headStyle.setFillForegroundColor((short) 9);
        return headStyle;
    }

}
