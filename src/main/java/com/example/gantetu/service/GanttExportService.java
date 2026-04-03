package com.example.gantetu.service;

import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import com.example.gantetu.excel.GanttExcelExporter;
import com.example.gantetu.excel.GanttHeaderStyleConfig;
import com.example.gantetu.service.query.GanttDetailQueryService;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.io.OutputStream;
import java.util.List;

/**
 * 甘特图导出服务。
 * <p>
 * 作为业务层门面，负责串联数据查询与 Excel 导出，不让控制器直接接触导出细节。
 */
@Service
@RequiredArgsConstructor
public class GanttExportService {

    /**
     * 负责真正写出 Excel 内容的导出器。
     */
    private final GanttExcelExporter exporter = new GanttExcelExporter();

    private final GanttDetailQueryService ganttDetailQueryService;

    /**
     * 将给定明细直接导出到指定输出流。
     *
     * @param outputStream 输出流，通常来自 HTTP 响应
     * @param title 甘特图标题
     * @param details 甘特图步骤明细
     * @param styleConfig 表头样式配置；为空时使用默认值
     */
    public void export(OutputStream outputStream,
                       String title,
                       List<GanttChartOfWellProgressDetail> details,
                       GanttHeaderStyleConfig styleConfig) {
        exporter.export(outputStream, title, details,
                styleConfig == null ? GanttHeaderStyleConfig.defaults() : styleConfig);
    }

    /**
     * 根据井 ID 查询步骤明细并导出。
     *
     * @param outputStream 输出流
     * @param title 标题
     * @param wellId 井 ID
     * @param styleConfig 样式配置
     */
    public void exportByWellId(OutputStream outputStream,
                               String title,
                               Long wellId,
                               GanttHeaderStyleConfig styleConfig) {
        List<GanttChartOfWellProgressDetail> details = ganttDetailQueryService.listByWellId(wellId);
        export(outputStream, title, details, styleConfig);
    }
}
