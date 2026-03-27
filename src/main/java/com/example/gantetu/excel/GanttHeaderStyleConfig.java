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
     * 一级/二级表头字号。
     */
    @Builder.Default
    private short headerFontSize = 11;

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
     * 获取默认样式配置。
     *
     * @return 默认配置实例
     */
    public static GanttHeaderStyleConfig defaults() {
        return GanttHeaderStyleConfig.builder().build();
    }
}
