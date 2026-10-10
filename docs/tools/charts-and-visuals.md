# Charts & Visualizations

Tools for creating Excel charts, data bars, color scales, icon sets, and conditional formatting rules.

---

## 1. Charts

### `excel_add_chart`
Creates and embeds a chart from tabular data into a worksheet.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `dataStartCell` | ✅ | Data range start (e.g., `A1`) |
| `dataEndCell` | ✅ | Data range end (e.g., `B5`) |
| `chartType` | ✅ | `column`, `bar`, `line`, `pie`, `doughnut`, `area`, `scatter`, `radar` |
| `targetCell` | ✅ | Placement top-left cell (e.g., `D2`) |
| `title` | ❌ | Chart title string |
| `width` | ❌ | Chart width in pixels (default: `480`) |
| `height` | ❌ | Chart height in pixels (default: `320`) |

### `excel_list_charts`
Lists all charts embedded in a worksheet.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |

### `excel_update_chart`
Updates data source range for an existing chart by index.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `chartIndex` | ✅ | 0-based chart index |
| `dataStartCell` | ✅ | New data start cell |
| `dataEndCell` | ✅ | New data end cell |

### `excel_delete_chart`
Deletes a chart by its 0-based index.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `chartIndex` | ✅ | 0-based chart index |

---

## 2. Conditional Formatting

### `excel_add_conditional_format`
Adds a rule-based conditional format to a cell range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `ruleType` | ✅ | `cellValue`, `containsText`, `blanks`, `errors` |
| `operator` | ❌ | `greaterThan`, `lessThan`, `equal`, `between`, etc. |
| `formula1` | ❌ | Primary comparison value or threshold |
| `formula2` | ❌ | Secondary comparison value (for `between`) |
| `format` | ✅ | Object with `backgroundColor`, `fontColor`, `bold`, `italic` |

### `excel_remove_conditional_format`
Removes conditional formatting rules affecting the specified range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |

---

## 3. Data Bars, Color Scales & Icon Sets

### `excel_add_data_bar`
Adds a gradient data bar visual fill proportional to cell values in the range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `color` | ❌ | Bar hex color (default: `638EC6`) |

### `excel_add_color_scale`
Applies a 2-color or 3-color heat-map scale across cell values.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `minColor` | ❌ | Minimum value color hex (default: `F8696B`) |
| `midColor` | ❌ | Midpoint color hex (optional 3-color scale) |
| `maxColor` | ❌ | Maximum value color hex (default: `63BE7B`) |

### `excel_add_icon_set`
Adds graphical rating or status icons (arrows, traffic lights, ratings) to numbers.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `iconSet` | ✅ | `3Arrows`, `3TrafficLights`, `4Arrows`, `4TrafficLights`, `5Arrows`, `5Rating` |

---

← [Back to index](README.md)
