package com.example.gantetu.dto;

import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

import java.util.Date;

/**
 * 单个甘特图阶段的导出 DTO。
 * <p>
 * 一个阶段同时携带计划与实际两个时间段，导出时会拆成 `Plan` 和 `Actual` 两行。
 */
@Data
public class GanttChartOfWellProgressDetail {

    /**
     * 阶段名称，例如钻井、完井、测试。
     */
    @Schema(description = "阶段")
    private String phase;

    /**
     * 计划开始日期。
     */
    @Schema(description = "计划开始时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date planStartDate;

    /**
     * 计划结束日期。
     */
    @Schema(description = "计划结束时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date planEndDate;

    /**
     * 实际开始日期。
     */
    @Schema(description = "实际开始时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date actualStartDate;

    /**
     * 实际结束日期。
     */
    @Schema(description = "实际结束时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date actualEndDate;
}
