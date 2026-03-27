package com.example.gantetu.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.handler.CellWriteHandler;
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
import java.util.*;

public class GanttExcelExporter {

    private static final int BASE_INFO_COLUMN_COUNT = 4;
    private static final int HEADER_ROWS = 3;
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");

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
                .registerWriteHandler(new GanttSheetHandler(details, timelineContext))
                .registerWriteHandler(new GanttCellHandler(details, timelineContext, styleConfig))
                .sheet("Gantt")
                .doWrite(rows);
    }

    private List<List<String>> buildRows(String title,
                                         List<GanttChartOfWellProgressDetail> details,
                                         TimelineContext timelineContext) {
        int totalColumns = BASE_INFO_COLUMN_COUNT + timelineContext.totalDays;
        List<List<String>> rows = new ArrayList<>();

        List<String> row0 = blankRow(totalColumns);
        row0.set(0, title);
        rows.add(row0);

        List<String> row1 = blankRow(totalColumns);
        row1.set(0, "Phase");
        row1.set(1, "Type");
        row1.set(2, "Start");
        row1.set(3, "End");
        for (MonthSpan monthSpan : timelineContext.monthSpans) {
            row1.set(BASE_INFO_COLUMN_COUNT + monthSpan.startOffset, monthSpan.label);
        }
        rows.add(row1);

        List<String> row2 = blankRow(totalColumns);
        for (int i = 0; i < timelineContext.totalDays; i++) {
            row2.set(BASE_INFO_COLUMN_COUNT + i, String.valueOf(timelineContext.start.plusDays(i).getDayOfMonth()));
        }
        rows.add(row2);

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

    private String format(Date date) {
        return date == null ? "" : DATE_FORMAT.format(date);
    }

    private List<String> blankRow(int totalColumns) {
        List<String> row = new ArrayList<>(totalColumns);
        for (int i = 0; i < totalColumns; i++) {
            row.add("");
        }
        return row;
    }

    private TimelineContext buildTimeline(List<GanttChartOfWellProgressDetail> details) {
        LocalDate minDate = null;
        LocalDate maxDate = null;

        for (GanttChartOfWellProgressDetail detail : details) {
            for (Date date : List.of(detail.getPlanStartDate(), detail.getPlanEndDate(), detail.getActualStartDate(), detail.getActualEndDate())) {
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
            LocalDate monthEnd = cursor.atEndOfMonth().isAfter(timelineEnd) ? timelineEnd : cursor.atEndOfMonth();
            int offset = (int) (monthStart.toEpochDay() - timelineStart.toEpochDay());
            int endOffset = (int) (monthEnd.toEpochDay() - timelineStart.toEpochDay());
            monthSpans.add(new MonthSpan(cursor.toString(), offset, endOffset));
            cursor = cursor.plusMonths(1);
        }

        return new TimelineContext(timelineStart, timelineEnd, days, monthSpans);
    }

    private record MonthSpan(String label, int startOffset, int endOffset) {
    }

    private record TimelineContext(LocalDate start, LocalDate end, int totalDays, List<MonthSpan> monthSpans) {
    }

    private static final class GanttSheetHandler implements SheetWriteHandler {
        private final List<GanttChartOfWellProgressDetail> details;
        private final TimelineContext timelineContext;

        private GanttSheetHandler(List<GanttChartOfWellProgressDetail> details, TimelineContext timelineContext) {
            this.details = details;
            this.timelineContext = timelineContext;
        }

        @Override
        public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
            Sheet sheet = writeSheetHolder.getSheet();
            int lastColumn = BASE_INFO_COLUMN_COUNT + timelineContext.totalDays - 1;
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastColumn));
            for (int col = 0; col < BASE_INFO_COLUMN_COUNT; col++) {
                sheet.addMergedRegion(new CellRangeAddress(1, 2, col, col));
            }
            for (MonthSpan monthSpan : timelineContext.monthSpans) {
                sheet.addMergedRegion(new CellRangeAddress(
                        1,
                        1,
                        BASE_INFO_COLUMN_COUNT + monthSpan.startOffset,
                        BASE_INFO_COLUMN_COUNT + monthSpan.endOffset));
            }
            for (int i = 0; i < details.size(); i++) {
                int startRow = HEADER_ROWS + i * 2;
                sheet.addMergedRegion(new CellRangeAddress(startRow, startRow + 1, 0, 0));
            }
            sheet.createFreezePane(BASE_INFO_COLUMN_COUNT, HEADER_ROWS);
            sheet.setColumnWidth(0, 22 * 256);
            sheet.setColumnWidth(1, 10 * 256);
            sheet.setColumnWidth(2, 14 * 256);
            sheet.setColumnWidth(3, 14 * 256);
            for (int i = BASE_INFO_COLUMN_COUNT; i <= lastColumn; i++) {
                sheet.setColumnWidth(i, 4 * 256);
            }
        }
    }

    private static final class GanttCellHandler implements CellWriteHandler {

        private final Map<Integer, RowTypeRange> rowRangeMap = new HashMap<>();
        private final TimelineContext timelineContext;
        private final GanttHeaderStyleConfig styleConfig;

        private CellStyle titleStyle;
        private CellStyle headerStyle;
        private CellStyle dayHeaderStyle;
        private CellStyle bodyStyle;
        private CellStyle planFillStyle;
        private CellStyle actualFillStyle;

        private GanttCellHandler(List<GanttChartOfWellProgressDetail> details,
                                 TimelineContext timelineContext,
                                 GanttHeaderStyleConfig styleConfig) {
            this.timelineContext = timelineContext;
            this.styleConfig = styleConfig;
            for (int i = 0; i < details.size(); i++) {
                int planRow = HEADER_ROWS + i * 2;
                int actualRow = planRow + 1;
                GanttChartOfWellProgressDetail detail = details.get(i);
                rowRangeMap.put(planRow, new RowTypeRange(true, toLocalDate(detail.getPlanStartDate()), toLocalDate(detail.getPlanEndDate())));
                rowRangeMap.put(actualRow, new RowTypeRange(false, toLocalDate(detail.getActualStartDate()), toLocalDate(detail.getActualEndDate())));
            }
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

            if (row == 0) {
                cell.setCellStyle(titleStyle);
                return;
            }
            if (row == 1) {
                cell.setCellStyle(headerStyle);
                return;
            }
            if (row == 2) {
                cell.setCellStyle(dayHeaderStyle);
                return;
            }

            if (col < BASE_INFO_COLUMN_COUNT) {
                cell.setCellStyle(bodyStyle);
                return;
            }

            RowTypeRange rowTypeRange = rowRangeMap.get(row);
            if (rowTypeRange == null || rowTypeRange.start == null || rowTypeRange.end == null) {
                cell.setCellStyle(bodyStyle);
                return;
            }
            LocalDate thisDate = timelineContext.start.plusDays(col - BASE_INFO_COLUMN_COUNT);
            boolean inRange = !thisDate.isBefore(rowTypeRange.start) && !thisDate.isAfter(rowTypeRange.end);
            if (inRange) {
                cell.setCellStyle(rowTypeRange.plan ? planFillStyle : actualFillStyle);
            } else {
                cell.setCellStyle(bodyStyle);
            }
        }

        private void initStylesIfNeeded(Workbook workbook) {
            if (titleStyle != null) {
                return;
            }
            titleStyle = createBaseStyle(workbook, styleConfig.getTitleBgColor(), styleConfig.getTitleFontColor(), styleConfig.getTitleFontSize(), true);
            headerStyle = createBaseStyle(workbook, styleConfig.getHeaderBgColor(), styleConfig.getHeaderFontColor(), styleConfig.getHeaderFontSize(), true);
            dayHeaderStyle = createBaseStyle(workbook, styleConfig.getHeaderBgColor(), styleConfig.getHeaderFontColor(), styleConfig.getHeaderFontSize(), false);
            bodyStyle = createBaseStyle(workbook, IndexedColors.WHITE.getIndex(), IndexedColors.BLACK.getIndex(), (short) 10, false);
            planFillStyle = createBaseStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), IndexedColors.BLACK.getIndex(), (short) 10, false);
            actualFillStyle = createBaseStyle(workbook, IndexedColors.ROSE.getIndex(), IndexedColors.BLACK.getIndex(), (short) 10, false);
        }

        private CellStyle createBaseStyle(Workbook workbook, short bgColor, short fontColor, short fontSize, boolean bold) {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(bgColor);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
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

        private LocalDate toLocalDate(Date date) {
            if (date == null) {
                return null;
            }
            return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        }

        private record RowTypeRange(boolean plan, LocalDate start, LocalDate end) {
        }
    }
}
