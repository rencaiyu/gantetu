package com.example.gantetu.dto;

import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

import java.util.Date;

@Data
public class GanttChartOfWellProgressDetail {
    @Schema(description = "步骤")
    private String phase;

    @Schema(description = "计划开始时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date planStartDate;

    @Schema(description = "计划结束时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date planEndDate;

    @Schema(description = "实际开始时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date actualStartDate;

    @Schema(description = "实际结束时间")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date actualEndDate;
}
