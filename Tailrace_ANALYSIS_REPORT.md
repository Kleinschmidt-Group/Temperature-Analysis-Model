# Tailrace Water Temperature Analysis — Review Report

*Generated 2026-07-06 12:57:14*

Primary station: **GCGW**  |  Summer months: [6, 7, 8]

This bundle summarizes the data, methods, and results of the analysis for review. The full source code is included alongside as `Tailrace_analysis_code_snapshot.py`.

## 1. Data Sources

**Water temperature (hourly, DART):**

- `GCGW` — Grand Coulee Tailrace (GCGW): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\GCGW\GCGW_Hourly_Data_1995_2025.xlsx`
- `FDRW` — Grand Coulee Forebay (FDRW): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FDRW\FDRW_Hourly_Data_1995_2025.xlsx`
- `CHJ` — Chief Joseph Forebay (CHJ): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\CHJ\CHJ_Hourly_Data_1995_2025.xlsx`
- `CHQW` — Chief Joseph Tailrace (CHQW): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\CHQW\CHQW_Hourly_Data_1995_2025.xlsx`

**Air temperature (USBR AgriMet daily):** `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\DailyClimatologicalData.csv`

**CMIP5 projection deltas (MACAv2-METDATA):**
- Tmax: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FutureTimeSeriesMax.csv`
- Tmin: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FutureTimeSeriesMin.csv`

## 2. Data Used (after QA/QC + date filter)

| Series | Label | Records | Start | End |
| --- | --- | --- | --- | --- |
| GCGW | Grand Coulee Tailrace (GCGW) | 241,409 | 1995 | 2026 |
| FDRW | Grand Coulee Forebay (FDRW) | 227,949 | 1997 | 2026 |
| CHJ | Chief Joseph Forebay (CHJ) | 175,971 | 1995 | 2026 |
| CHQW | Chief Joseph Tailrace (CHQW) | 146,852 | 1997 | 2025 |
| AIR | Air Temperature (AgriMet daily mean = (min+max)/2) | 8,818 | 2002 | 2026 |

## 3. QA/QC Flag Counts

**GCGW** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\GCGW\GCGW_Hourly_Data_1995_2025.xlsx`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 12638 |
| IQR Outliers (3.0xIQR) | 0 |
| Constant Spans (24h+) | 212 |
| Impossible Jumps (>3C/hr) | 20 |
| Data Gaps | 110 |

**FDRW** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FDRW\FDRW_Hourly_Data_1995_2025.xlsx`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 14919 |
| IQR Outliers (3.0xIQR) | 0 |
| Constant Spans (24h+) | 409 |
| Impossible Jumps (>3C/hr) | 32 |
| Data Gaps | 97 |

**CHJ** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\CHJ\CHJ_Hourly_Data_1995_2025.xlsx`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 6109 |
| IQR Outliers (3.0xIQR) | 0 |
| Constant Spans (24h+) | 180 |
| Impossible Jumps (>3C/hr) | 6 |
| Data Gaps | 154 |

**CHQW** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\CHQW\CHQW_Hourly_Data_1995_2025.xlsx`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 3605 |
| IQR Outliers (3.0xIQR) | 0 |
| Constant Spans (24h+) | 63 |
| Impossible Jumps (>3C/hr) | 18 |
| Data Gaps | 65 |

**AIR** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\DailyClimatologicalData.csv`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 2 |
| IQR Outliers (3.0xIQR) | 0 |
| Data Gaps | 6 |

## 4. Methods

1. **Mann-Kendall** trend test on annual JJA means.
2. **Theil-Sen** slope estimator (95% CI).
3. **Mohseni (1998)** nonlinear logistic air-water regression:
   `Tw = mu + (alpha - mu) / (1 + exp(gamma * (beta - Ta)))`
4. **CMIP5 delta** projections (RCP 4.5 / 8.5, 2040s / 2080s).
5. **Warm-day** scenario at 90th-percentile (10% exceedance) air temp.

### Mohseni Parameters

| Model | mu | alpha | gamma | beta | NSE |
| --- | --- | --- | --- | --- | --- |
| All weeks | 5.324 | 17.536 | 0.2311 | 13.337 | 0.5465 |
| JJA only | 11.578 | 16.925 | 0.8089 | 20.438 | 0.4621 |

Projections use the **all_weeks** model.

### Trends (annual JJA)

| Series | Slope (°C/yr) | MK tau | MK p |
| --- | --- | --- | --- |
| Primary water (GCGW) | -0.0014 | -0.0237 | 0.8650 |
| Air | +0.0338 | 0.1800 | 0.2158 |

**Observed baseline:** air 23.08 °C, water 15.61 °C

## 5. Results

### Station Summary (JJA)

| Station | Code | n_years | Mean_JJA_C | Mean_JJA_F | TheilSen_slope_C_yr | TheilSen_CI_low | TheilSen_CI_high | MK_tau | MK_S | MK_Z | MK_p | Trend | Significant_p05 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Grand Coulee Tailrace (GCGW) | GCGW | 31 | 15.603 | 60.085 | -0.001 | -0.037 | 0.025 | -0.024 | -11 | -0.170 | 0.865 | decreasing | No |
| Grand Coulee Forebay (FDRW) | FDRW | 29 | 16.476 | 61.658 | 0.029 | -0.016 | 0.075 | 0.163 | 66 | 1.219 | 0.223 | increasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | 31 | 15.919 | 60.654 | 0.010 | -0.019 | 0.031 | 0.080 | 37 | 0.612 | 0.541 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | 29 | 15.658 | 60.184 | 0.010 | -0.021 | 0.038 | 0.123 | 50 | 0.919 | 0.358 | increasing | No |

### Monthly Summary

| Station | Code | Month | n_years | Mean_C | Mean_F | TheilSen_slope_C_yr | TheilSen_CI_low | TheilSen_CI_high | MK_tau | MK_S | MK_Z | MK_p | Trend | Significant_p05 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Grand Coulee Tailrace (GCGW) | GCGW | June | 31 | 12.779 | 55.003 | -0.013 | -0.044 | 0.010 | -0.136 | -63 | -1.054 | 0.292 | decreasing | No |
| Grand Coulee Tailrace (GCGW) | GCGW | July | 31 | 15.741 | 60.334 | -0.006 | -0.039 | 0.027 | -0.050 | -23 | -0.374 | 0.709 | decreasing | No |
| Grand Coulee Tailrace (GCGW) | GCGW | August | 31 | 18.293 | 64.927 | -0.011 | -0.042 | 0.019 | -0.092 | -43 | -0.714 | 0.475 | decreasing | No |
| Grand Coulee Forebay (FDRW) | FDRW | June | 29 | 13.168 | 55.702 | 0.031 | -0.026 | 0.081 | 0.148 | 60 | 1.107 | 0.268 | increasing | No |
| Grand Coulee Forebay (FDRW) | FDRW | July | 29 | 16.622 | 61.920 | 0.046 | -0.004 | 0.113 | 0.236 | 96 | 1.782 | 0.075 | increasing | No |
| Grand Coulee Forebay (FDRW) | FDRW | August | 29 | 19.624 | 67.323 | -0.002 | -0.054 | 0.054 | -0.012 | -5 | -0.075 | 0.940 | decreasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | June | 31 | 13.026 | 55.447 | -0.005 | -0.034 | 0.024 | -0.054 | -25 | -0.408 | 0.683 | decreasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | July | 31 | 16.108 | 60.994 | 0.008 | -0.020 | 0.034 | 0.088 | 41 | 0.680 | 0.497 | increasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | August | 31 | 18.568 | 65.423 | 0.008 | -0.018 | 0.032 | 0.097 | 45 | 0.748 | 0.455 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | June | 29 | 12.831 | 55.097 | 0.004 | -0.025 | 0.048 | 0.035 | 14 | 0.244 | 0.807 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | July | 29 | 15.786 | 60.415 | 0.014 | -0.021 | 0.047 | 0.118 | 48 | 0.882 | 0.378 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | August | 29 | 18.227 | 64.808 | 0.014 | -0.016 | 0.057 | 0.089 | 36 | 0.656 | 0.511 | increasing | No |

### Combined JJA Projections

| Scenario | Delta_Air_C | Proj_Air_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | 0.000 | 23.080 | 16.370 | 61.470 | 0.000 | 0.000 |
| 2040s RCP 4.5 | 2.360 | 25.440 | 16.830 | 62.300 | 0.460 | 0.830 |
| 2040s RCP 8.5 | 2.970 | 26.050 | 16.920 | 62.460 | 0.550 | 0.990 |
| 2080s RCP 4.5 | 3.460 | 26.540 | 16.980 | 62.570 | 0.610 | 1.100 |
| 2080s RCP 8.5 | 6.110 | 29.190 | 17.230 | 63.020 | 0.860 | 1.540 |

### June Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_C | Proj_Air_C | Obs_Baseline_Tw_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F | Obs_Max_Tw_C | Proj_Max_Tw_C | Proj_Max_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | June | 0.000 | 19.940 | 19.940 | 12.780 | 12.780 | 55.010 | 0.000 | 0.000 | 16.500 | 16.500 | 61.700 |
| 2040s RCP 4.5 | June | 2.360 | 19.940 | 22.300 | 12.780 | 13.600 | 56.470 | 0.810 | 1.470 | 16.500 | 17.310 | 63.170 |
| 2040s RCP 8.5 | June | 2.970 | 19.940 | 22.910 | 12.780 | 13.760 | 56.770 | 0.980 | 1.760 | 16.500 | 17.480 | 63.460 |
| 2080s RCP 4.5 | June | 3.460 | 19.940 | 23.400 | 12.780 | 13.880 | 56.980 | 1.090 | 1.970 | 16.500 | 17.590 | 63.670 |
| 2080s RCP 8.5 | June | 6.110 | 19.940 | 26.050 | 12.780 | 14.350 | 57.830 | 1.570 | 2.820 | 16.500 | 18.070 | 64.520 |

### July Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_C | Proj_Air_C | Obs_Baseline_Tw_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F | Obs_Max_Tw_C | Proj_Max_Tw_C | Proj_Max_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | July | 0.000 | 24.970 | 24.970 | 15.710 | 15.710 | 60.290 | 0.000 | 0.000 | 21.100 | 21.100 | 69.980 |
| 2040s RCP 4.5 | July | 2.360 | 24.970 | 27.330 | 15.710 | 16.030 | 60.850 | 0.310 | 0.570 | 21.100 | 21.410 | 70.550 |
| 2040s RCP 8.5 | July | 2.970 | 24.970 | 27.940 | 15.710 | 16.090 | 60.960 | 0.370 | 0.670 | 21.100 | 21.470 | 70.650 |
| 2080s RCP 4.5 | July | 3.460 | 24.970 | 28.430 | 15.710 | 16.130 | 61.030 | 0.420 | 0.750 | 21.100 | 21.520 | 70.730 |
| 2080s RCP 8.5 | July | 6.110 | 24.970 | 31.080 | 15.710 | 16.290 | 61.330 | 0.580 | 1.040 | 21.100 | 21.680 | 71.020 |

### August Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_C | Proj_Air_C | Obs_Baseline_Tw_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F | Obs_Max_Tw_C | Proj_Max_Tw_C | Proj_Max_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | August | 0.000 | 24.350 | 24.350 | 18.280 | 18.280 | 64.910 | 0.000 | 0.000 | 22.200 | 22.200 | 71.960 |
| 2040s RCP 4.5 | August | 2.360 | 24.350 | 26.710 | 18.280 | 18.640 | 65.550 | 0.360 | 0.640 | 22.200 | 22.560 | 72.600 |
| 2040s RCP 8.5 | August | 2.970 | 24.350 | 27.320 | 18.280 | 18.710 | 65.680 | 0.420 | 0.760 | 22.200 | 22.620 | 72.720 |
| 2080s RCP 4.5 | August | 3.460 | 24.350 | 27.810 | 18.280 | 18.760 | 65.760 | 0.470 | 0.850 | 22.200 | 22.670 | 72.810 |
| 2080s RCP 8.5 | August | 6.110 | 24.350 | 30.460 | 18.280 | 18.940 | 66.100 | 0.660 | 1.190 | 22.200 | 22.860 | 73.150 |

### June Warm-Day (90th pct) Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_P90_C | Proj_Air_P90_C | Obs_Baseline_Tw_P90_C | Proj_Tw_P90_C | Proj_Tw_P90_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | June | 0.000 | 24.900 | 24.900 | 14.300 | 14.300 | 57.740 | 0.000 | 0.000 |
| 2040s RCP 4.5 | June | 2.360 | 24.900 | 27.260 | 14.300 | 14.620 | 58.310 | 0.320 | 0.570 |
| 2040s RCP 8.5 | June | 2.970 | 24.900 | 27.870 | 14.300 | 14.680 | 58.420 | 0.380 | 0.680 |
| 2080s RCP 4.5 | June | 3.460 | 24.900 | 28.360 | 14.300 | 14.720 | 58.500 | 0.420 | 0.760 |
| 2080s RCP 8.5 | June | 6.110 | 24.900 | 31.010 | 14.300 | 14.890 | 58.800 | 0.590 | 1.060 |

### July Warm-Day (90th pct) Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_P90_C | Proj_Air_P90_C | Obs_Baseline_Tw_P90_C | Proj_Tw_P90_C | Proj_Tw_P90_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | July | 0.000 | 29.100 | 29.100 | 17.600 | 17.600 | 63.680 | 0.000 | 0.000 |
| 2040s RCP 4.5 | July | 2.360 | 29.100 | 31.460 | 17.600 | 17.730 | 63.910 | 0.130 | 0.230 |
| 2040s RCP 8.5 | July | 2.970 | 29.100 | 32.070 | 17.600 | 17.750 | 63.950 | 0.150 | 0.270 |
| 2080s RCP 4.5 | July | 3.460 | 29.100 | 32.560 | 17.600 | 17.770 | 63.980 | 0.170 | 0.300 |
| 2080s RCP 8.5 | July | 6.110 | 29.100 | 35.210 | 17.600 | 17.830 | 64.100 | 0.230 | 0.420 |

### August Warm-Day (90th pct) Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_P90_C | Proj_Air_P90_C | Obs_Baseline_Tw_P90_C | Proj_Tw_P90_C | Proj_Tw_P90_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1995–2026) | August | 0.000 | 28.360 | 28.360 | 19.500 | 19.500 | 67.100 | 0.000 | 0.000 |
| 2040s RCP 4.5 | August | 2.360 | 28.360 | 30.720 | 19.500 | 19.650 | 67.370 | 0.150 | 0.270 |
| 2040s RCP 8.5 | August | 2.970 | 28.360 | 31.330 | 19.500 | 19.680 | 67.420 | 0.180 | 0.320 |
| 2080s RCP 4.5 | August | 3.460 | 28.360 | 31.820 | 19.500 | 19.700 | 67.460 | 0.200 | 0.360 |
| 2080s RCP 8.5 | August | 6.110 | 28.360 | 34.470 | 19.500 | 19.780 | 67.600 | 0.280 | 0.500 |

### Weekly Water Temperature Stats

| week | mean_tw | min_tw | max_tw | n_years | primary_month |
| --- | --- | --- | --- | --- | --- |
| 22.000 | 11.445 | 10.064 | 13.280 | 26.000 | 6.000 |
| 23.000 | 11.955 | 10.694 | 13.520 | 31.000 | 6.000 |
| 24.000 | 12.555 | 9.887 | 13.785 | 31.000 | 6.000 |
| 25.000 | 13.260 | 12.302 | 14.762 | 31.000 | 6.000 |
| 26.000 | 14.014 | 12.269 | 15.481 | 47.000 | 6.000 |
| 27.000 | 14.523 | 12.948 | 15.955 | 41.000 | 7.000 |
| 28.000 | 15.370 | 13.213 | 19.587 | 31.000 | 7.000 |
| 29.000 | 16.067 | 13.657 | 18.977 | 31.000 | 7.000 |
| 30.000 | 16.697 | 14.131 | 18.694 | 33.000 | 7.000 |
| 31.000 | 17.287 | 14.472 | 19.272 | 55.000 | 8.000 |
| 32.000 | 17.933 | 15.500 | 19.625 | 31.000 | 8.000 |
| 33.000 | 18.388 | 16.646 | 20.523 | 31.000 | 8.000 |
| 34.000 | 18.718 | 17.019 | 20.620 | 31.000 | 8.000 |
| 35.000 | 18.956 | 17.652 | 21.179 | 30.000 | 8.000 |
| 36.000 | 19.478 | 18.096 | 21.704 | 5.000 | 8.000 |

## 6. Figures

| File | Description |
| --- | --- |
| Fig1_Annual_Summer_Trends.png | Annual Jun/Jul/Aug mean water temperatures with trends |
| Fig2_Seasonal_Climatology.png | Mean annual temperature cycle (water + air overlay) |
| Fig3_FDRW_Trend_Detail.png | Primary station detailed trend + air overlay |
| Fig4_JJA_Daily_By_Year.png | Daily JJA temperatures, all years overlaid |
| Fig5_Mohseni_Regression.png | Mohseni air-water logistic regression |
| Fig6_Mohseni_Observed_vs_Predicted.png | Observed vs predicted weekly water temps |
| Fig7_JJA_Obs_vs_Pred_TimeSeries.png | JJA annual observed vs predicted |
| Fig8_Mohseni_All_Years_Overlaid.png | All years overlaid on Mohseni curve |
| Fig9_Projected_Temps_Bar.png | Projected temperatures bar chart |
| Fig10_Projection_Points_Mohseni.png | Projection points on Mohseni curve |
| Fig11_Period_Comparison_Boxplots.png | Early vs recent period comparison |
| Fig12_Monthly_Projections_Bar.png | Monthly projections (Jun/Jul/Aug) |
| Fig13_Monthly_Projection_Points.png | Monthly projections on curve |
| Fig14_Observed_Weekly_MinMaxMean.png | Observed weekly temperature ranges |
| Fig15_Projected_Weekly_MinMaxMean.png | Projected weekly temperature ranges |
| Fig16_Average_vs_WarmDay_Comparison.png | Average day vs warm day comparison |
