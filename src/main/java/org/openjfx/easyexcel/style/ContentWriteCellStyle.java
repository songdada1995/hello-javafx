package org.openjfx.easyexcel.style;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import lombok.Data;
import org.openjfx.constants.EasyExcelConstants;

/**
 * @author songbo
 * @version 1.0
 * @date 2023/5/29 19:17
 */
@Data
public class ContentWriteCellStyle extends WriteCellStyle {

    /**
     * 列宽
     */
    private Integer columnWidth;

    /**
     * 行高
     */
    private Short rowHeight;

    /**
     * 默认内容样式
     *
     * @return
     */
    public static ContentWriteCellStyle getDefaultContentStyle() {
        ContentWriteCellStyle contentStyle = new ContentWriteCellStyle();
        WriteFont writeFont = new WriteFont();
        writeFont.setBold(false);
        writeFont.setFontName(EasyExcelConstants.DefaultConstants.DEFAULT_FONT_NAME);
        writeFont.setFontHeightInPoints(EasyExcelConstants.DefaultConstants.DEFAULT_FONT_SIZE);
        contentStyle.setWriteFont(writeFont);
        contentStyle.setColumnWidth(EasyExcelConstants.DefaultConstants.DEFAULT_COLUMN_WIDTH);
        contentStyle.setRowHeight(EasyExcelConstants.DefaultConstants.DEFAULT_ROW_HEIGHT);
        contentStyle.setFillForegroundColor((short) 9);
        return contentStyle;
    }

}
