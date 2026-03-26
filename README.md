# SpringBoot + EasyExcel 甘特图导出示例

## 版本
- Spring Boot `3.3.5`
- EasyExcel `3.3.2`

## 导出能力
- 第 1 行：总标题（整行合并）
- 第 2~3 行：
  - `A2-A3` 合并：`Phase`
  - `B2-B3` 合并：`Type`
  - `C2-C3` 合并：`Start`
  - `D2-D3` 合并：`End`
  - `E2~`：按月合并显示（如 `2026-03`）
  - `E3~`：逐天显示（日）
- 数据区：每个 phase 占两行
  - 第 1 行 `Plan`
  - 第 2 行 `Actual`
  - phase 名称列纵向合并（例如 `A4-A5`）
- 时间轴上：
  - 计划区间填充蓝色
  - 实际区间填充红色
- 前两行标题字体大小/字体颜色/背景色支持自定义（请求参数可选）

## 接口
`POST /api/gantt/export`

请求体示例：

```json
{
  "title": "Well Progress Gantt Chart",
  "titleFontSize": 16,
  "titleFontColor": "WHITE",
  "titleBgColor": "DARK_BLUE",
  "headerFontSize": 11,
  "headerFontColor": "WHITE",
  "headerBgColor": "GREY_50_PERCENT",
  "details": [
    {
      "phase": "Drilling",
      "planStartDate": "2026-03-01",
      "planEndDate": "2026-03-15",
      "actualStartDate": "2026-03-03",
      "actualEndDate": "2026-03-20"
    }
  ]
}
```

> 颜色请使用 `IndexedColors` 枚举名（如 `RED`、`DARK_BLUE`、`GREY_50_PERCENT`）。

