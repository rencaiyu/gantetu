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
import java.util.List;

/**
 * 甘特图导出控制器。
 * <p>
 * 提供 HTTP 接口接收导出参数，构建样式配置，并将生成的 Excel 作为附件流式返回。
 */
@RestController
@RequestMapping("/api/gantt")
@RequiredArgsConstructor
public class GanttExportController {

    /**
     * 导出服务。
     */
    private final GanttExportService ganttExportService;

    /**
     * 导出甘特图 Excel。
     *
     * @param request  导出请求（标题、井 ID、颜色与字号配置）
     * @param response HTTP 响应对象
     * @throws IOException 写出响应流时可能抛出 IO 异常
     */
    @PostMapping("/export")
    public void export(@RequestBody GanttExportRequest request,
                       HttpServletResponse response) throws IOException {
        if (request.getWellId() == null) {
            throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "wellId 不能为空");
        }

        String fileName = URLEncoder.encode("gantt-chart", StandardCharsets.UTF_8)
                .replaceAll("\\+", "%20");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition",
                "attachment;filename*=utf-8''" + fileName + ".xlsx");

        GanttHeaderStyleConfig style = GanttHeaderStyleConfig.builder()
                .titleFontSize(request.getTitleFontSize() == null ? 16 : request.getTitleFontSize())
                .titleBgColor(parseColor(request.getTitleBgColor(), IndexedColors.DARK_BLUE.getIndex()))
                .titleFontColor(parseColor(request.getTitleFontColor(), IndexedColors.WHITE.getIndex()))
                .titleRowHeight(request.getTitleRowHeight() == null ? 28f : request.getTitleRowHeight())
                .headerFontSize(request.getHeaderFontSize() == null ? 11 : request.getHeaderFontSize())
                .levelOneHeaderRowHeight(request.getLevelOneHeaderRowHeight() == null ? 22f : request.getLevelOneHeaderRowHeight())
                .levelTwoHeaderRowHeight(request.getLevelTwoHeaderRowHeight() == null ? 18f : request.getLevelTwoHeaderRowHeight())
                .headerBgColor(parseColor(request.getHeaderBgColor(), IndexedColors.GREY_50_PERCENT.getIndex()))
                .headerFontColor(parseColor(request.getHeaderFontColor(), IndexedColors.WHITE.getIndex()))
                .levelOneHeaderWidth(toExcelColumnWidth(request.getLevelOneHeaderWidth(), 14))
                .levelTwoHeaderWidth(toExcelColumnWidth(request.getLevelTwoHeaderWidth(), 4))
                .build();

        ganttExportService.exportByWellId(response.getOutputStream(), request.getTitle(), request.getWellId(), style);
    }

    /**
     * 将字符串颜色名转换为 Apache POI 颜色索引。
     * <p>
     * 示例输入：DARK_BLUE、WHITE、GREY_50_PERCENT。
     *
     * @param colorName    颜色名称
     * @param defaultColor 默认颜色索引（当 colorName 为空或非法时使用）
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
     * 导出请求体。
     */
    @Data
    public static class GanttExportRequest {

        /** 甘特图标题。 */
        private String title = "Well Progress Gantt Chart";

        /** 井 ID（用于从数据库查询甘特图步骤明细）。 */
        private Long wellId;

        /** 兼容历史接口保留字段（已废弃，不再用于导出）。 */
        @Deprecated
        private List<GanttChartOfWellProgressDetail> details;

        /** 标题字号（可选）。 */
        private Short titleFontSize;

        /** 标题字体颜色（可选，IndexedColors 名称）。 */
        private String titleFontColor;

        /** 标题背景颜色（可选，IndexedColors 名称）。 */
        private String titleBgColor;

        /** 标题行高度（可选，单位 point）。 */
        private Float titleRowHeight;

        /** 表头字号（可选）。 */
        private Short headerFontSize;

        /** 一级表头行高（可选，单位 point，作用于月份/字段名行）。 */
        private Float levelOneHeaderRowHeight;

        /** 二级表头行高（可选，单位 point，作用于日号行）。 */
        private Float levelTwoHeaderRowHeight;

        /** 表头字体颜色（可选，IndexedColors 名称）。 */
        private String headerFontColor;

        /** 表头背景颜色（可选，IndexedColors 名称）。 */
        private String headerBgColor;

        /** 一级表头列宽（可选，单位字符，作用于 Phase/Type/Start/End 列）。 */
        private Integer levelOneHeaderWidth;

        /** 二级表头列宽（可选，单位字符，作用于时间轴日列）。 */
        private Integer levelTwoHeaderWidth;
    }
}
