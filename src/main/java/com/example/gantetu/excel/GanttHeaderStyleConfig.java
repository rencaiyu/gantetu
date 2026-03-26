package com.example.gantetu.excel;

import lombok.Builder;
import lombok.Getter;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Builder
public class GanttHeaderStyleConfig {
    @Builder.Default
    private short titleFontSize = 16;
    @Builder.Default
    private short titleFontColor = IndexedColors.WHITE.getIndex();
    @Builder.Default
    private short titleBgColor = IndexedColors.DARK_BLUE.getIndex();

    @Builder.Default
    private short headerFontSize = 11;
    @Builder.Default
    private short headerFontColor = IndexedColors.WHITE.getIndex();
    @Builder.Default
    private short headerBgColor = IndexedColors.GREY_50_PERCENT.getIndex();

    public static GanttHeaderStyleConfig defaults() {
        return GanttHeaderStyleConfig.builder().build();
    }
}
