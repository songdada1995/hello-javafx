package org.openjfx.easyexcel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.BooleanUtils;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.openjfx.easyexcel.style.HeadWriteCellStyle;
import org.openjfx.exception.BasicException;

import java.lang.reflect.Field;
import java.util.*;

/**
 * @author songbo
 * @date 2023-05-29
 */
public class CustomStyleWriteHandler implements CellWriteHandler {

    /**
     * 复杂表头自定义样式数组
     */
    private List<ComplexHeadStyleVO> headStyleList;

    /**
     * 内容样式
     */
    private List<ComplexContentStyleVO> contentStyleList;

    public CustomStyleWriteHandler(List<ComplexHeadStyleVO> headStyleList, List<ComplexContentStyleVO> contentStyleList) {
        this.headStyleList = headStyleList;
        this.contentStyleList = contentStyleList;
    }

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        // 设置头部样式
        if (BooleanUtils.isTrue(context.getHead())) {
            // 第一个单元格
            // 只要不是头 一定会有数据 当然fill的情况 可能要context.getCellDataList() ,这个需要看模板，因为一个单元格会有多个 WriteCellData
            WriteCellData<?> cellData = context.getFirstCellData();
            // 这里需要去cellData 获取样式
            // 很重要的一个原因是 WriteCellStyle 和 dataFormatData绑定的 简单的说 比如你加了 DateTimeFormat
            // ，已经将writeCellStyle里面的dataFormatData 改了 如果你自己new了一个WriteCellStyle，可能注解的样式就失效了
            // 然后 getOrCreateStyle 用于返回一个样式，如果为空，则创建一个后返回
            WriteCellStyle writeCellStyle = cellData.getOrCreateStyle();

            if (CollectionUtils.isNotEmpty(headStyleList)) {
                headStyleList.forEach(complexHeadStyleVO -> {
                    HeadWriteCellStyle headWriteCellStyle = complexHeadStyleVO.getHeadWriteCellStyle();
                    Cell cell = context.getCell();
                    Integer relativeRowIndex = context.getRelativeRowIndex();
                    // 取出队列中的自定义表头信息，与当前坐标比较，判断是否相符
                    if (cell.getColumnIndex() == complexHeadStyleVO.getY() && relativeRowIndex.equals(complexHeadStyleVO.getX())) {
                        // 设置自定义的表头颜色
                        if (null != headWriteCellStyle.getFillForegroundColor()) {
                            writeCellStyle.setFillForegroundColor(headWriteCellStyle.getFillForegroundColor());
                        }
                        // 设置字体格式
                        if (null != headWriteCellStyle.getWriteFont()) {
                            writeCellStyle.setWriteFont(headWriteCellStyle.getWriteFont());
                        }
                        //  设置内容位置
                        if (null != headWriteCellStyle.getHorizontalAlignment()) {
                            writeCellStyle.setHorizontalAlignment(headWriteCellStyle.getHorizontalAlignment());
                        }
                        if (null != headWriteCellStyle.getVerticalAlignment()) {
                            writeCellStyle.setVerticalAlignment(headWriteCellStyle.getVerticalAlignment());
                        }
                        //  设置列宽，行高
                        Sheet sheet = cell.getSheet();
                        if (null != headWriteCellStyle.getColumnWidth()) {
                            sheet.setColumnWidth(complexHeadStyleVO.getY(), headWriteCellStyle.getColumnWidth());
                        }
                        if (null != headWriteCellStyle.getRowHeight()) {
                            sheet.getRow(complexHeadStyleVO.getX()).setHeight(headWriteCellStyle.getRowHeight());
                        }
                    }
                });
            }
        }

        // 也可设置内容样式
    }

    /**
     * 构建自定义写策略
     *
     * @param clazz
     * @return
     */
    public static CustomStyleWriteHandler buildDefaultWriteHandler(Class clazz) {
        List<Field> fieldList = new ArrayList<>();
        while (clazz != null) {
            fieldList.addAll(Arrays.asList(clazz.getDeclaredFields()));
            clazz = clazz.getSuperclass();
        }

        Set<Integer> indexSet = new HashSet<>();
        for (Field field : fieldList) {
            if (field.isAnnotationPresent(ExcelProperty.class)) {
                ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
                indexSet.add(annotation.index());
            }
        }

        if (CollectionUtils.isEmpty(indexSet)) {
            throw new BasicException("请使用@ExcelProperty注解");
        }

        ArrayList<ComplexHeadStyleVO> headStyleList = new ArrayList<>();
        for (int i = 0; i < indexSet.size(); i++) {
            headStyleList.add(new ComplexHeadStyleVO(0, i, HeadWriteCellStyle.getDefaultHeadStyle()));
        }
        return new CustomStyleWriteHandler(headStyleList, null);
    }

}