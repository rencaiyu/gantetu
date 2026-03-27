package com.example.gantetu.dto;

import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

import java.util.Date;

/**
 * 甘特图单个步骤的数据传输对象。
 * <p>
 * 每个步骤同时包含计划时间段与实际时间段，导出时会渲染为两行（Plan / Actual）。
 */
@Data
public class GanttChartOfWellProgressDetail {

    /**
     * 步骤名称（例如：钻井、完井、测试等）。
     */
    @Schema(description = "步骤")
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
