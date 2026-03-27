CREATE TABLE IF NOT EXISTS gantt_step_detail
(
    id                BIGINT PRIMARY KEY AUTO_INCREMENT COMMENT '主键',
    well_id           BIGINT       NOT NULL COMMENT '井ID',
    phase             VARCHAR(128) NOT NULL COMMENT '步骤名称',
    plan_start_date   DATE         NOT NULL COMMENT '计划开始时间',
    plan_end_date     DATE         NOT NULL COMMENT '计划结束时间',
    actual_start_date DATE         NULL COMMENT '实际开始时间',
    actual_end_date   DATE         NULL COMMENT '实际结束时间',
    sort_order        INT          NOT NULL DEFAULT 0 COMMENT '排序号',
    INDEX idx_well_id_sort (well_id, sort_order)
) ENGINE = InnoDB
  DEFAULT CHARSET = utf8mb4
  COMMENT ='甘特图步骤明细表';
