package com.example.gantetu.entity;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

import java.util.Date;

/**
 * 甘特图步骤明细表实体。
 */
@Data
@TableName("gantt_step_detail")
public class GanttStepDetailEntity {

    /** 主键。 */
    @TableId(type = IdType.AUTO)
    private Long id;

    /** 井 ID。 */
    private Long wellId;

    /** 步骤名称。 */
    private String phase;

    /** 计划开始时间。 */
    private Date planStartDate;

    /** 计划结束时间。 */
    private Date planEndDate;

    /** 实际开始时间。 */
    private Date actualStartDate;

    /** 实际结束时间。 */
    private Date actualEndDate;

    /** 排序号，越小越靠前。 */
    private Integer sortOrder;
}
