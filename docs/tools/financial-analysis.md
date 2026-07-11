# Financial Analysis

Tools for period comparison, variance analysis, balance checks, and anomaly detection.

## `excel_period_comparison`

Compare current vs previous period values.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `currentValueRange` | ✅ | Current period value range |
| `previousValueRange` | ✅ | Previous period value range |
| `outputRange` | ✅ | Output range for results |
| `showPercentage` | ❌ | Show percentage change |

## `excel_variance_analysis`

Calculate budget vs actual variance.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `budgetRange` | ✅ | Budget value range |
| `actualRange` | ✅ | Actual value range |
| `outputRange` | ✅ | Output range for results |
| `showPercentage` | ❌ | Show percentage variance |

## `excel_check_balance`

Verify that debits equal credits (accounting balance check).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `debitRange` | ✅ | Debit values range |
| `creditRange` | ✅ | Credit values range |

## `excel_find_anomalies`

Identify outliers using standard deviation.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `rangeStart` | ✅ | Range start cell |
| `rangeEnd` | ✅ | Range end cell |
| `stdDevThreshold` | ❌ | Standard deviation threshold for anomaly detection |

---

← [Back to index](README.md)