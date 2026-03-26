package com.example.gantetu.controller;

import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import com.example.gantetu.excel.GanttHeaderStyleConfig;
import com.example.gantetu.service.GanttExportService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.Data;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

@RestController
@RequestMapping("/api/gantt")
@RequiredArgsConstructor
public class GanttExportController {

    private final GanttExportService ganttExportService;

    @PostMapping("/export")
    public void export(@RequestBody GanttExportRequest request, HttpServletResponse response) throws IOException {
        String fileName = URLEncoder.encode("gantt-chart", StandardCharsets.UTF_8).replaceAll("\\+", "%20");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

        GanttHeaderStyleConfig style = GanttHeaderStyleConfig.builder()
                .titleFontSize(request.getTitleFontSize() == null ? 16 : request.getTitleFontSize())
                .titleBgColor(parseColor(request.getTitleBgColor(), IndexedColors.DARK_BLUE.getIndex()))
                .titleFontColor(parseColor(request.getTitleFontColor(), IndexedColors.WHITE.getIndex()))
                .headerFontSize(request.getHeaderFontSize() == null ? 11 : request.getHeaderFontSize())
                .headerBgColor(parseColor(request.getHeaderBgColor(), IndexedColors.GREY_50_PERCENT.getIndex()))
                .headerFontColor(parseColor(request.getHeaderFontColor(), IndexedColors.WHITE.getIndex()))
                .build();

        ganttExportService.export(response.getOutputStream(), request.getTitle(), request.getDetails(), style);
    }

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

    @Data
    public static class GanttExportRequest {
        private String title = "Well Progress Gantt Chart";
        private List<GanttChartOfWellProgressDetail> details;

        // 可选：IndexedColors 名称，如 DARK_BLUE、WHITE、GREY_50_PERCENT
        private Short titleFontSize;
        private String titleFontColor;
        private String titleBgColor;
        private Short headerFontSize;
        private String headerFontColor;
        private String headerBgColor;
    }
}
