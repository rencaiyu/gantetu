package com.example.gantetu.controller;

import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import com.example.gantetu.excel.GanttHeaderStyleConfig;
import com.example.gantetu.service.GanttExportService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.Data;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.server.ResponseStatusException;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * 甘特图导出控制器。
 * <p>
 * 对外暴露 HTTP 接口，负责接收导出参数、组装样式配置，并将生成的 Excel 以附件形式返回。
 */
@RestController
@RequestMapping("/api/gantt")
@RequiredArgsConstructor
public class GanttExportController {

    private static final DateTimeFormatter FILE_TS_FORMAT = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");
    private static final DateTimeFormatter TITLE_TS_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss'Z'");

    /**
     * 业务导出服务。
     */
    private final GanttExportService ganttExportService;

    /**
     * 导出甘特图 Excel。
     *
     * @param request 导出请求，包含标题、井 ID 和样式参数
     * @param response HTTP 响应对象
     * @throws IOException 写出响应流时可能抛出的异常
     */
    @PostMapping("/export")
    public void export(@RequestBody GanttExportRequest request,
                       HttpServletResponse response) throws IOException {
        if (request.getWellId() == null) {
            throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "wellId 不能为空");
        }

        LocalDateTime exportTime = LocalDateTime.now(ZoneOffset.UTC);
        String timestamp = exportTime.format(FILE_TS_FORMAT);
        String fileName = URLEncoder.encode("gantt-chart-" + timestamp, StandardCharsets.UTF_8)
                .replaceAll("\\+", "%20");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition",
                "attachment;filename*=utf-8''" + fileName + ".xlsx");

        GanttHeaderStyleConfig style = GanttHeaderStyleConfig.builder()
                .titleFontSize(request.getTitleFontSize() == null ? 16 : request.getTitleFontSize())
                .titleBgColor(parseColor(request.getTitleBgColor(), IndexedColors.DARK_BLUE.getIndex()))
                .titleFontColor(parseColor(request.getTitleFontColor(), IndexedColors.WHITE.getIndex()))
                .titleRowHeight(normalizeRowHeight(request.getTitleRowHeight(), 28f))
                .headerFontSize(request.getHeaderFontSize() == null ? 11 : request.getHeaderFontSize())
                .levelOneHeaderRowHeight(normalizeRowHeight(request.getLevelOneHeaderRowHeight(), 22f))
                .levelTwoHeaderRowHeight(normalizeRowHeight(request.getLevelTwoHeaderRowHeight(), 18f))
                .headerBgColor(parseColor(request.getHeaderBgColor(), IndexedColors.GREY_50_PERCENT.getIndex()))
                .headerFontColor(parseColor(request.getHeaderFontColor(), IndexedColors.WHITE.getIndex()))
                .levelOneHeaderWidth(toExcelColumnWidth(request.getLevelOneHeaderWidth(), 14))
                .levelTwoHeaderWidth(toExcelColumnWidth(request.getLevelTwoHeaderWidth(), 4))
                .build();

        String exportTitle = request.getTitle() + " (generated at " + exportTime.format(TITLE_TS_FORMAT) + ")";
        ganttExportService.exportByWellId(response.getOutputStream(), exportTitle, request.getWellId(), style);
    }

    /**
     * 将颜色名称转换为 Apache POI 颜色索引。
     *
     * @param colorName 颜色名称，例如 `DARK_BLUE`
     * @param defaultColor 默认颜色索引
     * @return 颜色索引
     */
    private short parseColor(String colorName, short defaultColor) {
        if (colorName == null || colorName.isBlank()) {
            return defaultColor;
        }
        try {
            return IndexedColors.valueOf(colorName.trim().toUpperCase()).getIndex();
        } catch (IllegalArgumentException ex) {
            return defaultColor;
        }
    }

    /**
     * 将前端传入的字符宽度转换为 Excel 列宽单位。
     */
    private int toExcelColumnWidth(Integer widthInChars, int defaultChars) {
        int chars = (widthInChars == null || widthInChars <= 0) ? defaultChars : widthInChars;
        return chars * 256;
    }

    /**
     * 规范化行高入参，统一输出为 point。
     * <p>
     * 兼容两类输入：
     * 1. 直接传 point，例如 `28`
     * 2. 传 twips，例如 `560`，会自动换算成 `28pt`
     */
    private float normalizeRowHeight(Float rowHeight, float defaultPoint) {
        if (rowHeight == null || rowHeight <= 0) {
            return defaultPoint;
        }
        float excelMaxPoint = 409.5f;
        if (rowHeight > excelMaxPoint) {
            return rowHeight / 20f;
        }
        if (rowHeight >= 100f && isAlmostInteger(rowHeight) && ((int) Math.round(rowHeight)) % 10 == 0) {
            return rowHeight / 20f;
        }
        return rowHeight;
    }

    /**
     * 判断浮点值是否足够接近整数，避免 twips 识别时受精度误差影响。
     */
    private boolean isAlmostInteger(float value) {
        return Math.abs(value - Math.round(value)) < 0.0001f;
    }

    /**
     * 导出请求体。
     */
    @Data
    public static class GanttExportRequest {

        /** 甘特图标题。 */
        private String title = "Well Progress Gantt Chart";

        /** 井 ID，用于从数据库查询步骤明细。 */
        private Long wellId;

        /** 兼容历史接口保留字段，已废弃，不再参与导出。 */
        @Deprecated
        private List<GanttChartOfWellProgressDetail> details;

        /** 标题字号。 */
        private Short titleFontSize;

        /** 标题字体颜色，使用 `IndexedColors` 名称。 */
        private String titleFontColor;

        /** 标题背景颜色，使用 `IndexedColors` 名称。 */
        private String titleBgColor;

        /** 标题行高，支持 point，也兼容 twips。 */
        private Float titleRowHeight;

        /** 表头字号。 */
        private Short headerFontSize;

        /** 一级表头行高，作用于月份行和固定字段行。 */
        private Float levelOneHeaderRowHeight;

        /** 二级表头行高，作用于日号行。 */
        private Float levelTwoHeaderRowHeight;

        /** 表头字体颜色，使用 `IndexedColors` 名称。 */
        private String headerFontColor;

        /** 表头背景颜色，使用 `IndexedColors` 名称。 */
        private String headerBgColor;

        /** 一级表头列宽，作用于 `Phase/Type/Start/End`。 */
        private Integer levelOneHeaderWidth;

        /** 二级表头列宽，作用于时间轴日列。 */
        private Integer levelTwoHeaderWidth;
    }
}
