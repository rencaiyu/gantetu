package com.example.gantetu.entity;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

import java.util.Date;

/**
 * `gantt_step_detail` 表对应的实体。
 */
@Data
@TableName("gantt_step_detail")
public class GanttStepDetailEntity {

    /** 主键。 */
    @TableId(type = IdType.AUTO)
    private Long id;

    /** 井 ID，用于隔离不同井的数据。 */
    private Long wellId;

    /** 阶段名称，例如钻井、完井、测试。 */
    private String phase;

    /** 计划开始时间。 */
    private Date planStartDate;

    /** 计划结束时间。 */
    private Date planEndDate;

    /** 实际开始时间，可为空。 */
    private Date actualStartDate;

    /** 实际结束时间，可为空。 */
    private Date actualEndDate;

    /** 排序号，值越小越靠前。 */
    private Integer sortOrder;
}
