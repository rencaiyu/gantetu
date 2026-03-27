# SpringBoot + EasyExcel 甘特图导出示例

## 版本
- Spring Boot `3.3.5`
- EasyExcel `3.3.2`
- MyBatis-Plus `3.5.5`
- MySQL `8.x`

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
- 标题行高度支持自定义（请求参数可选）
- 一级/二级表头列宽支持自定义（请求参数可选）

## MySQL + MyBatis-Plus 集成

### 1) 配置数据库连接
在 `src/main/resources/application.yml` 修改你的数据库连接：

```yaml
spring:
  datasource:
    url: jdbc:mysql://127.0.0.1:3306/gantetu?useUnicode=true&characterEncoding=UTF-8&useSSL=false&serverTimezone=UTC
    username: root
    password: root
```

### 2) 初始化表结构
执行 `src/main/resources/db/schema.sql`：

```sql
CREATE TABLE IF NOT EXISTS gantt_step_detail (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  well_id BIGINT NOT NULL,
  phase VARCHAR(128) NOT NULL,
  plan_start_date DATE NOT NULL,
  plan_end_date DATE NOT NULL,
  actual_start_date DATE NULL,
  actual_end_date DATE NULL,
  sort_order INT NOT NULL DEFAULT 0
);
```

## 接口
`POST /api/gantt/export`

> 现在导出时会根据 `wellId` 从 MySQL 查询步骤明细列表，不再依赖前端传 `details`。

请求体示例：

```json
{
  "title": "Well Progress Gantt Chart",
  "wellId": 10001,
  "titleFontSize": 16,
  "titleFontColor": "WHITE",
  "titleBgColor": "DARK_BLUE",
  "titleRowHeight": 28,
  "headerFontSize": 11,
  "headerFontColor": "WHITE",
  "headerBgColor": "GREY_50_PERCENT",
  "levelOneHeaderWidth": 14,
  "levelTwoHeaderWidth": 4
}
```

> 颜色请使用 `IndexedColors` 枚举名（如 `RED`、`DARK_BLUE`、`GREY_50_PERCENT`）。
>
> `titleRowHeight` 单位为 point；`levelOneHeaderWidth`、`levelTwoHeaderWidth` 单位为“字符宽度”。
