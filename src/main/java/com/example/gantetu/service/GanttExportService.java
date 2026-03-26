package com.example.gantetu.service;

import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import com.example.gantetu.excel.GanttExcelExporter;
import com.example.gantetu.excel.GanttHeaderStyleConfig;
import org.springframework.stereotype.Service;

import java.io.OutputStream;
import java.util.List;

@Service
public class GanttExportService {

    private final GanttExcelExporter exporter = new GanttExcelExporter();

    public void export(OutputStream outputStream,
                       String title,
                       List<GanttChartOfWellProgressDetail> details,
                       GanttHeaderStyleConfig styleConfig) {
        exporter.export(outputStream, title, details, styleConfig == null ? GanttHeaderStyleConfig.defaults() : styleConfig);
    }
}
