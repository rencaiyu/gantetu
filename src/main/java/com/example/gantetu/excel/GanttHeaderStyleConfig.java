package com.example.gantetu.excel;

import lombok.Builder;
import lombok.Getter;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 甘特图表头样式配置。
 * <p>
 * 该配置用于控制标题与表头的字号、字体颜色、背景色。
 */
@Getter
@Builder
public class GanttHeaderStyleConfig {

    /**
     * Excel 宽度单位换算常量（1 个字符宽度 = 256 单位）。
     */
    private static final int EXCEL_WIDTH_UNIT = 256;

    /**
     * 标题字号。
     */
    @Builder.Default
    private short titleFontSize = 16;

    /**
     * 标题字体颜色（IndexedColors 索引值）。
     */
    @Builder.Default
    private short titleFontColor = IndexedColors.WHITE.getIndex();

    /**
     * 标题背景色（IndexedColors 索引值）。
     */
    @Builder.Default
    private short titleBgColor = IndexedColors.DARK_BLUE.getIndex();

    /**
     * 标题行高度（point）。
     */
    @Builder.Default
    private float titleRowHeight = 28f;

    /**
     * 一级/二级表头字号。
     */
    @Builder.Default
    private short headerFontSize = 11;

    /**
     * 一级表头行高（point，对应月份/字段名行）。
     */
    @Builder.Default
    private float levelOneHeaderRowHeight = 22f;

    /**
     * 二级表头行高（point，对应日号行）。
     */
    @Builder.Default
    private float levelTwoHeaderRowHeight = 18f;

    /**
     * 一级/二级表头字体颜色（IndexedColors 索引值）。
     */
    @Builder.Default
    private short headerFontColor = IndexedColors.WHITE.getIndex();

    /**
     * 一级/二级表头背景色（IndexedColors 索引值）。
     */
    @Builder.Default
    private short headerBgColor = IndexedColors.GREY_50_PERCENT.getIndex();

    /**
     * 一级表头列宽（字符数，应用于 Phase/Type/Start/End 列）。
     */
    @Builder.Default
    private int levelOneHeaderWidth = 14 * EXCEL_WIDTH_UNIT;

    /**
     * 二级表头列宽（字符数，应用于时间轴日列）。
     */
    @Builder.Default
    private int levelTwoHeaderWidth = 4 * EXCEL_WIDTH_UNIT;

    /**
     * 获取默认样式配置。
     *
     * @return 默认配置实例
     */
    public static GanttHeaderStyleConfig defaults() {
        return GanttHeaderStyleConfig.builder().build();
    }
}
