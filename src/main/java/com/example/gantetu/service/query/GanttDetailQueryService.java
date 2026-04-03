package com.example.gantetu.service.query;

import com.baomidou.mybatisplus.core.conditions.query.LambdaQueryWrapper;
import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import com.example.gantetu.entity.GanttStepDetailEntity;
import com.example.gantetu.mapper.GanttStepDetailMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * 甘特图步骤查询服务。
 */
@Service
@RequiredArgsConstructor
public class GanttDetailQueryService {

    private final GanttStepDetailMapper ganttStepDetailMapper;

    /**
     * 根据井 ID 查询甘特图步骤明细，并转换为导出 DTO。
     *
     * @param wellId 井 ID
     * @return 导出 DTO 列表
     */
    public List<GanttChartOfWellProgressDetail> listByWellId(Long wellId) {
        LambdaQueryWrapper<GanttStepDetailEntity> queryWrapper = new LambdaQueryWrapper<GanttStepDetailEntity>()
                .eq(GanttStepDetailEntity::getWellId, wellId)
                .orderByAsc(GanttStepDetailEntity::getSortOrder)
                .orderByAsc(GanttStepDetailEntity::getId);

        return ganttStepDetailMapper.selectList(queryWrapper)
                .stream()
                .map(this::toDto)
                .toList();
    }

    /**
     * 将数据库实体转换为导出使用的 DTO，避免导出层依赖持久化对象。
     */
    private GanttChartOfWellProgressDetail toDto(GanttStepDetailEntity entity) {
        GanttChartOfWellProgressDetail dto = new GanttChartOfWellProgressDetail();
        dto.setPhase(entity.getPhase());
        dto.setPlanStartDate(entity.getPlanStartDate());
        dto.setPlanEndDate(entity.getPlanEndDate());
        dto.setActualStartDate(entity.getActualStartDate());
        dto.setActualEndDate(entity.getActualEndDate());
        return dto;
    }
}
