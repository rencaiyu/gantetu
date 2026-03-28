package com.example.gantetu.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.example.gantetu.dto.GanttChartOfWellProgressDetail;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 甘特图 Excel 导出核心实现。
 * <p>
 * 该类负责：
 * 1) 根据计划/实际日期自动构建时间轴；
 * 2) 组装导出的二维表格数据；
 * 3) 通过 EasyExcel 的 Sheet/Cell 回调完成单元格合并、样式与区间填色。
 */
public class GanttExcelExporter {

    /** 基础信息列数：Phase、Type、Start、End。 */
    private static final int BASE_INFO_COLUMN_COUNT = 4;

    /** 表头行数：标题行 + 月份行 + 天行。 */
    private static final int HEADER_ROWS = 3;

    /** 日期展示格式。 */
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");

    /**
     * 执行甘特图导出。
     *
     * @param outputStream 输出流
     * @param title        标题
     * @param details      业务明细数据
     * @param styleConfig  样式配置
     */
    public void export(OutputStream outputStream,
                       String title,
                       List<GanttChartOfWellProgressDetail> details,
                       GanttHeaderStyleConfig styleConfig) {
        if (details == null || details.isEmpty()) {
            throw new IllegalArgumentException("details 不能为空");
        }

        TimelineContext timelineContext = buildTimeline(details);
        List<List<String>> rows = buildRows(title, details, timelineContext);

        EasyExcel.write(outputStream)
                .needHead(false)
                .registerWriteHandler(new GanttSheetHandler(details, timelineContext, styleConfig))
                .registerWriteHandler(new GanttRowHeightHandler(styleConfig))
                .registerWriteHandler(new GanttCellHandler(styleConfig))
                .sheet("Gantt")
                .doWrite(rows);
    }

    /**
     * 组装写入 Excel 的内容行。
     */
    private List<List<String>> buildRows(String title,
                                         List<GanttChartOfWellProgressDetail> details,
                                         TimelineContext timelineContext) {
        int totalColumns = BASE_INFO_COLUMN_COUNT + timelineContext.totalDays;
        List<List<String>> rows = new ArrayList<>();

        // 第 0 行：总标题
        List<String> row0 = blankRow(totalColumns);
        row0.set(0, title);
        rows.add(row0);

        // 第 1 行：固定字段 + 月份标签
        List<String> row1 = blankRow(totalColumns);
        row1.set(0, "Phase");
        row1.set(1, "Type");
        row1.set(2, "Start");
        row1.set(3, "End");
        for (MonthSpan monthSpan : timelineContext.monthSpans) {
            row1.set(BASE_INFO_COLUMN_COUNT + monthSpan.startOffset, monthSpan.label);
        }
        rows.add(row1);

        // 第 2 行：每天的日号（1~31）
        List<String> row2 = blankRow(totalColumns);
        for (int i = 0; i < timelineContext.totalDays; i++) {
            row2.set(BASE_INFO_COLUMN_COUNT + i,
                    String.valueOf(timelineContext.start.plusDays(i).getDayOfMonth()));
        }
        rows.add(row2);

        // 从第 3 行开始，每个阶段两行：计划行 + 实际行
        for (GanttChartOfWellProgressDetail detail : details) {
            List<String> planRow = blankRow(totalColumns);
            planRow.set(0, detail.getPhase());
            planRow.set(1, "Plan");
            planRow.set(2, format(detail.getPlanStartDate()));
            planRow.set(3, format(detail.getPlanEndDate()));
            rows.add(planRow);

            List<String> actualRow = blankRow(totalColumns);
            actualRow.set(1, "Actual");
            actualRow.set(2, format(detail.getActualStartDate()));
            actualRow.set(3, format(detail.getActualEndDate()));
            rows.add(actualRow);
        }
        return rows;
    }

    /** 将日期格式化为 yyyy-MM-dd 字符串。 */
    private String format(Date date) {
        return date == null ? "" : DATE_FORMAT.format(date);
    }

    /**
     * 构造空白行，长度等于总列数。
     */
    private List<String> blankRow(int totalColumns) {
        List<String> row = new ArrayList<>(totalColumns);
        for (int i = 0; i < totalColumns; i++) {
            row.add("");
        }
        return row;
    }

    /**
     * 从业务明细中计算时间轴。
     * <p>
     * 时间轴规则：
     * - 起点取最早日期所在月的 1 号；
     * - 终点取最晚日期所在月的最后一天；
     * - 生成每月区间（偏移起止）供表头合并使用。
     */
    private TimelineContext buildTimeline(List<GanttChartOfWellProgressDetail> details) {
        LocalDate minDate = null;
        LocalDate maxDate = null;

        for (GanttChartOfWellProgressDetail detail : details) {
            Date[] dates = {
                    detail.getPlanStartDate(),
                    detail.getPlanEndDate(),
                    detail.getActualStartDate(),
                    detail.getActualEndDate()
            };
            for (Date date : dates) {
                if (date == null) {
                    continue;
                }
                LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                minDate = (minDate == null || localDate.isBefore(minDate)) ? localDate : minDate;
                maxDate = (maxDate == null || localDate.isAfter(maxDate)) ? localDate : maxDate;
            }
        }
        if (minDate == null || maxDate == null) {
            throw new IllegalArgumentException("计划/实际时间不能为空");
        }

        LocalDate timelineStart = minDate.withDayOfMonth(1);
        LocalDate timelineEnd = maxDate.withDayOfMonth(maxDate.lengthOfMonth());
        int days = (int) (timelineEnd.toEpochDay() - timelineStart.toEpochDay() + 1);

        List<MonthSpan> monthSpans = new ArrayList<>();
        YearMonth cursor = YearMonth.from(timelineStart);
        YearMonth last = YearMonth.from(timelineEnd);
        while (!cursor.isAfter(last)) {
            LocalDate first = cursor.atDay(1);
            LocalDate monthStart = first.isBefore(timelineStart) ? timelineStart : first;
            LocalDate monthEnd = cursor.atEndOfMonth().isAfter(timelineEnd)
                    ? timelineEnd
                    : cursor.atEndOfMonth();
            int offset = (int) (monthStart.toEpochDay() - timelineStart.toEpochDay());
            int endOffset = (int) (monthEnd.toEpochDay() - timelineStart.toEpochDay());
            monthSpans.add(new MonthSpan(cursor.toString(), offset, endOffset));
            cursor = cursor.plusMonths(1);
        }

        return new TimelineContext(timelineStart, timelineEnd, days, monthSpans);
    }

    /**
     * 单个月份在时间轴中的跨度信息。
     *
     * @param label       月份标签（yyyy-MM）
     * @param startOffset 月份起点相对于时间轴起点的偏移
     * @param endOffset   月份终点相对于时间轴起点的偏移
     */
    private record MonthSpan(String label, int startOffset, int endOffset) {
    }

    /**
     * 时间轴上下文。
     *
     * @param start     轴起点
     * @param end       轴终点
     * @param totalDays 总天数
     * @param monthSpans 月度跨度集合
     */
    private record TimelineContext(LocalDate start, LocalDate end, int totalDays, List<MonthSpan> monthSpans) {
    }

    /**
     * Sheet 级处理器：负责合并单元格、冻结窗格、设置列宽。
     */
    private static final class GanttSheetHandler implements SheetWriteHandler {
        private final List<GanttChartOfWellProgressDetail> details;
        private final TimelineContext timelineContext;
        private final GanttHeaderStyleConfig styleConfig;

        private GanttSheetHandler(List<GanttChartOfWellProgressDetail> details,
                                  TimelineContext timelineContext,
                                  GanttHeaderStyleConfig styleConfig) {
            this.details = details;
            this.timelineContext = timelineContext;
            this.styleConfig = styleConfig;
        }

        @Override
        public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder,
                                     WriteSheetHolder writeSheetHolder) {
            Sheet sheet = writeSheetHolder.getSheet();
            int lastColumn = BASE_INFO_COLUMN_COUNT + timelineContext.totalDays - 1;
            int firstDataRow = HEADER_ROWS + 1;
            int lastDataRow = HEADER_ROWS + details.size() * 2;

            // 标题横跨全列
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastColumn));

            // 前 4 列头（Phase/Type/Start/End）在第 1~2 行纵向合并
            for (int col = 0; col < BASE_INFO_COLUMN_COUNT; col++) {
                sheet.addMergedRegion(new CellRangeAddress(1, 2, col, col));
            }

            // 月份行按月份跨度横向合并
            for (MonthSpan monthSpan : timelineContext.monthSpans) {
                sheet.addMergedRegion(new CellRangeAddress(
                        1,
                        1,
                        BASE_INFO_COLUMN_COUNT + monthSpan.startOffset,
                        BASE_INFO_COLUMN_COUNT + monthSpan.endOffset));
            }

            // 每个阶段的 Plan/Actual 两行，Phase 列纵向合并
            for (int i = 0; i < details.size(); i++) {
                int startRow = HEADER_ROWS + i * 2;
                sheet.addMergedRegion(new CellRangeAddress(startRow, startRow + 1, 0, 0));
            }

            // 冻结基础信息列与表头
            sheet.createFreezePane(BASE_INFO_COLUMN_COUNT, HEADER_ROWS);

            // 一级/二级表头列宽
            for (int i = 0; i < BASE_INFO_COLUMN_COUNT; i++) {
                sheet.setColumnWidth(i, styleConfig.getLevelOneHeaderWidth());
            }
            for (int i = BASE_INFO_COLUMN_COUNT; i <= lastColumn; i++) {
                sheet.setColumnWidth(i, styleConfig.getLevelTwoHeaderWidth());
            }

            // 数据区动态条件格式：
            // 当用户在 Excel 中调整 C(Start)/D(End) 时，时间轴颜色会自动刷新。
            if (!details.isEmpty() && timelineContext.totalDays > 0) {
                applyDynamicRangeColorRules(sheet, firstDataRow, lastDataRow, lastColumn);
            }
        }

        /**
         * 使用条件格式实现动态区间着色（Plan 蓝色，Actual 红色）。
         * <p>
         * 规则基于每一行 C/D 列的日期值与时间轴列对应日期的比较，在 Excel 中修改后会自动重算颜色。
         */
        private void applyDynamicRangeColorRules(Sheet sheet,
                                                 int firstDataRow,
                                                 int lastDataRow,
                                                 int lastColumn) {
            SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

            String timelineStartExpr = String.format(
                    "DATE(%d,%d,%d)+COLUMN()-%d",
                    timelineContext.start.getYear(),
                    timelineContext.start.getMonthValue(),
                    timelineContext.start.getDayOfMonth(),
                    BASE_INFO_COLUMN_COUNT + 1
            );
            String startExpr = "IF(ISNUMBER($C4),$C4,DATEVALUE($C4))";
            String endExpr = "IF(ISNUMBER($D4),$D4,DATEVALUE($D4))";

            String planFormula = String.format(
                    "AND($B4=\"Plan\",%s>=%s,%s<=%s)",
                    timelineStartExpr, startExpr, timelineStartExpr, endExpr
            );
            ConditionalFormattingRule planRule = scf.createConditionalFormattingRule(planFormula);
            PatternFormatting planFormatting = planRule.createPatternFormatting();
            planFormatting.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            planFormatting.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());

            String actualFormula = String.format(
                    "AND($B4=\"Actual\",%s>=%s,%s<=%s)",
                    timelineStartExpr, startExpr, timelineStartExpr, endExpr
            );
            ConditionalFormattingRule actualRule = scf.createConditionalFormattingRule(actualFormula);
            PatternFormatting actualFormatting = actualRule.createPatternFormatting();
            actualFormatting.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            actualFormatting.setFillForegroundColor(IndexedColors.ROSE.getIndex());

            CellRangeAddress[] regions = {
                    new CellRangeAddress(firstDataRow - 1, lastDataRow - 1, BASE_INFO_COLUMN_COUNT, lastColumn)
            };
            scf.addConditionalFormatting(regions, planRule, actualRule);
        }
    }

    /**
     * 行级处理器：在行创建完成后设置表头行高，避免被后续写入过程覆盖。
     */
    private static final class GanttRowHeightHandler implements RowWriteHandler {

        private final GanttHeaderStyleConfig styleConfig;

        private GanttRowHeightHandler(GanttHeaderStyleConfig styleConfig) {
            this.styleConfig = styleConfig;
        }

        @Override
        public void afterRowDispose(WriteSheetHolder writeSheetHolder,
                                    WriteTableHolder writeTableHolder,
                                    Row row,
                                    Integer relativeRowIndex,
                                    Boolean isHead) {
            if (row == null) {
                return;
            }
            int rowIndex = row.getRowNum();
            if (rowIndex == 0) {
                row.setHeightInPoints(styleConfig.getTitleRowHeight());
                return;
            }
            if (rowIndex == 1) {
                row.setHeightInPoints(styleConfig.getLevelOneHeaderRowHeight());
                return;
            }
            if (rowIndex == 2) {
                row.setHeightInPoints(styleConfig.getLevelTwoHeaderRowHeight());
            }
        }
    }

    /**
     * Cell 级处理器：按行/列位置应用样式，并对日期区间进行填色。
     */
    private static final class GanttCellHandler implements CellWriteHandler {

        /** 颜色/字号配置。 */
        private final GanttHeaderStyleConfig styleConfig;

        private CellStyle titleStyle;
        private CellStyle headerStyle;
        private CellStyle dayHeaderStyle;
        private CellStyle bodyStyle;
        private GanttCellHandler(GanttHeaderStyleConfig styleConfig) {
            this.styleConfig = styleConfig;
        }

        @Override
        public void afterCellDispose(WriteSheetHolder writeSheetHolder,
                                     WriteTableHolder writeTableHolder,
                                     List<com.alibaba.excel.metadata.data.WriteCellData<?>> list,
                                     Cell cell,
                                     com.alibaba.excel.metadata.Head head,
                                     Integer integer,
                                     Boolean isHead) {
            initStylesIfNeeded(cell.getSheet().getWorkbook());
            int row = cell.getRowIndex();
            int col = cell.getColumnIndex();

            // 标题行样式
            if (row == 0) {
                cell.setCellStyle(titleStyle);
                return;
            }
            // 月份/字段名行样式
            if (row == 1) {
                cell.setCellStyle(headerStyle);
                return;
            }
            // 天数行样式
            if (row == 2) {
                cell.setCellStyle(dayHeaderStyle);
                return;
            }

            // 左侧基础信息列
            if (col < BASE_INFO_COLUMN_COUNT) {
                cell.setCellStyle(bodyStyle);
                return;
            }

            cell.setCellStyle(bodyStyle);
        }

        /** 延迟初始化样式，避免重复创建对象。 */
        private void initStylesIfNeeded(Workbook workbook) {
            if (titleStyle != null) {
                return;
            }
            titleStyle = createBaseStyle(workbook,
                    styleConfig.getTitleBgColor(),
                    styleConfig.getTitleFontColor(),
                    styleConfig.getTitleFontSize(),
                    true);
            headerStyle = createBaseStyle(workbook,
                    styleConfig.getHeaderBgColor(),
                    styleConfig.getHeaderFontColor(),
                    styleConfig.getHeaderFontSize(),
                    true);
            dayHeaderStyle = createBaseStyle(workbook,
                    styleConfig.getHeaderBgColor(),
                    styleConfig.getHeaderFontColor(),
                    styleConfig.getHeaderFontSize(),
                    false);
            bodyStyle = createBaseStyle(workbook,
                    IndexedColors.WHITE.getIndex(),
                    IndexedColors.BLACK.getIndex(),
                    (short) 10,
                    false);
        }

        /**
         * 创建统一基础样式（边框、对齐、背景、字体）。
         */
        private CellStyle createBaseStyle(Workbook workbook,
                                          short bgColor,
                                          short fontColor,
                                          short fontSize,
                                          boolean bold) {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(bgColor);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            //水平居中
//            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            //左对齐
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);

            Font font = workbook.createFont();
            font.setColor(fontColor);
            font.setFontHeightInPoints(fontSize);
            font.setBold(bold);
            cellStyle.setFont(font);
            return cellStyle;
        }

    }
}
