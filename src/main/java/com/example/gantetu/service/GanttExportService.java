package com.example.gantetu.service;

import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import com.example.gantetu.excel.GanttExcelExporter;
import com.example.gantetu.excel.GanttHeaderStyleConfig;
import org.springframework.stereotype.Service;

import java.io.OutputStream;
import java.util.List;

/**
 * 甘特图导出服务。
 * <p>
 * 该类作为业务层门面，屏蔽控制层对底层 Excel 生成细节的感知。
 */
@Service
public class GanttExportService {

    /**
     * 实际执行 Excel 构建与写出的导出器。
     */
    private final GanttExcelExporter exporter = new GanttExcelExporter();

    /**
     * 导出甘特图到指定输出流。
     *
     * @param outputStream 输出流（通常来自 HTTP 响应）
     * @param title        甘特图标题
     * @param details      甘特图步骤数据列表
     * @param styleConfig  表头样式配置；当传入为空时自动使用默认样式
     */
    public void export(OutputStream outputStream,
                       String title,
                       List<GanttChartOfWellProgressDetail> details,
                       GanttHeaderStyleConfig styleConfig) {
        exporter.export(outputStream, title, details,
                styleConfig == null ? GanttHeaderStyleConfig.defaults() : styleConfig);
    }
}
