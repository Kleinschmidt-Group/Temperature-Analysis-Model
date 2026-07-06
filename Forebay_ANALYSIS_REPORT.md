# Forebay Water Temperature Analysis — Review Report

*Generated 2026-07-06 12:52:19*

Primary station: **FDRW**  |  Summer months: [6, 7, 8]

This bundle summarizes the data, methods, and results of the analysis for review. The full source code is included alongside as `Forebay_analysis_code_snapshot.py`.

## 1. Data Sources

**Water temperature (hourly, DART):**

- `FDRW` — Grand Coulee Forebay (FDRW): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FDRW\FDRW_Hourly_Data_1995_2025.xlsx`
- `GCGW` — Grand Coulee Tailrace (GCGW): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\GCGW\GCGW_Hourly_Data_1995_2025.xlsx`
- `CHJ` — Chief Joseph Forebay (CHJ): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\CHJ\CHJ_Hourly_Data_1995_2025.xlsx`
- `CHQW` — Chief Joseph Tailrace (CHQW): `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\CHQW\CHQW_Hourly_Data_1995_2025.xlsx`

**Air temperature (USBR AgriMet daily):** `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\DailyClimatologicalData.csv`

**CMIP5 projection deltas (MACAv2-METDATA):**
- Tmax: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FutureTimeSeriesMax.csv`
- Tmin: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FutureTimeSeriesMin.csv`

## 2. Data Used (after QA/QC + date filter)

| Series | Label | Records | Start | End |
| --- | --- | --- | --- | --- |
| FDRW | Grand Coulee Forebay (FDRW) | 227,949 | 1997 | 2026 |
| GCGW | Grand Coulee Tailrace (GCGW) | 241,409 | 1995 | 2026 |
| CHJ | Chief Joseph Forebay (CHJ) | 175,971 | 1995 | 2026 |
| CHQW | Chief Joseph Tailrace (CHQW) | 146,852 | 1997 | 2025 |
| AIR | Air Temperature (AgriMet daily mean = (min+max)/2) | 8,818 | 2002 | 2026 |

## 3. QA/QC Flag Counts

**FDRW** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\FDRW\FDRW_Hourly_Data_1995_2025.xlsx`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 14919 |
| IQR Outliers (3.0xIQR) | 0 |
| Constant Spans (24h+) | 409 |
| Impossible Jumps (>3C/hr) | 32 |
| Data Gaps | 97 |

**GCGW** (source: `C:\Users\Ethan.Muhlestein\OneDrive - Kleinschmidt Associates\Documents\CJD_Temp_Monitoring\Data\GCGW\GCGW_Hourly_Data_1995_2025.xlsx`)
| Check | Flagged |
| --- | --- |
| Physical Bounds | 12638 |
| IQR Outliers (3.0xIQR) | 0 |
| Constant Spans (24h+) | 212 |
| Impossible Jumps (>3C/hr) | 20 |
| Data Gaps | 110 |

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
| All weeks | 5.370 | 18.829 | 0.2360 | 13.706 | 0.5803 |
| JJA only | 11.981 | 18.271 | 0.7814 | 20.577 | 0.4502 |

Projections use the **all_weeks** model.

### Trends (annual JJA)

| Series | Slope (°C/yr) | MK tau | MK p |
| --- | --- | --- | --- |
| Primary water (FDRW) | +0.0287 | 0.1626 | 0.2227 |
| Air | +0.0338 | 0.1800 | 0.2158 |

**Observed baseline:** air 23.08 °C, water 16.48 °C

## 5. Results

### Station Summary (JJA)

| Station | Code | n_years | Mean_JJA_C | Mean_JJA_F | TheilSen_slope_C_yr | TheilSen_CI_low | TheilSen_CI_high | MK_tau | MK_S | MK_Z | MK_p | Trend | Significant_p05 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Grand Coulee Forebay (FDRW) | FDRW | 29 | 16.476 | 61.658 | 0.029 | -0.016 | 0.075 | 0.163 | 66 | 1.219 | 0.223 | increasing | No |
| Grand Coulee Tailrace (GCGW) | GCGW | 31 | 15.603 | 60.085 | -0.001 | -0.037 | 0.025 | -0.024 | -11 | -0.170 | 0.865 | decreasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | 31 | 15.919 | 60.654 | 0.010 | -0.019 | 0.031 | 0.080 | 37 | 0.612 | 0.541 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | 29 | 15.658 | 60.184 | 0.010 | -0.021 | 0.038 | 0.123 | 50 | 0.919 | 0.358 | increasing | No |

### Monthly Summary

| Station | Code | Month | n_years | Mean_C | Mean_F | TheilSen_slope_C_yr | TheilSen_CI_low | TheilSen_CI_high | MK_tau | MK_S | MK_Z | MK_p | Trend | Significant_p05 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Grand Coulee Forebay (FDRW) | FDRW | June | 29 | 13.168 | 55.702 | 0.031 | -0.026 | 0.081 | 0.148 | 60 | 1.107 | 0.268 | increasing | No |
| Grand Coulee Forebay (FDRW) | FDRW | July | 29 | 16.622 | 61.920 | 0.046 | -0.004 | 0.113 | 0.236 | 96 | 1.782 | 0.075 | increasing | No |
| Grand Coulee Forebay (FDRW) | FDRW | August | 29 | 19.624 | 67.323 | -0.002 | -0.054 | 0.054 | -0.012 | -5 | -0.075 | 0.940 | decreasing | No |
| Grand Coulee Tailrace (GCGW) | GCGW | June | 31 | 12.779 | 55.003 | -0.013 | -0.044 | 0.010 | -0.136 | -63 | -1.054 | 0.292 | decreasing | No |
| Grand Coulee Tailrace (GCGW) | GCGW | July | 31 | 15.741 | 60.334 | -0.006 | -0.039 | 0.027 | -0.050 | -23 | -0.374 | 0.709 | decreasing | No |
| Grand Coulee Tailrace (GCGW) | GCGW | August | 31 | 18.293 | 64.927 | -0.011 | -0.042 | 0.019 | -0.092 | -43 | -0.714 | 0.475 | decreasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | June | 31 | 13.026 | 55.447 | -0.005 | -0.034 | 0.024 | -0.054 | -25 | -0.408 | 0.683 | decreasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | July | 31 | 16.108 | 60.994 | 0.008 | -0.020 | 0.034 | 0.088 | 41 | 0.680 | 0.497 | increasing | No |
| Chief Joseph Forebay (CHJ) | CHJ | August | 31 | 18.568 | 65.423 | 0.008 | -0.018 | 0.032 | 0.097 | 45 | 0.748 | 0.455 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | June | 29 | 12.831 | 55.097 | 0.004 | -0.025 | 0.048 | 0.035 | 14 | 0.244 | 0.807 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | July | 29 | 15.786 | 60.415 | 0.014 | -0.021 | 0.047 | 0.118 | 48 | 0.882 | 0.378 | increasing | No |
| Chief Joseph Tailrace (CHQW) | CHQW | August | 29 | 18.227 | 64.808 | 0.014 | -0.016 | 0.057 | 0.089 | 36 | 0.656 | 0.511 | increasing | No |

### Combined JJA Projections

| Scenario | Delta_Air_C | Proj_Air_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | 0.000 | 23.080 | 17.500 | 63.500 | 0.000 | 0.000 |
| 2040s RCP 4.5 | 2.360 | 25.440 | 18.030 | 64.460 | 0.530 | 0.960 |
| 2040s RCP 8.5 | 2.970 | 26.050 | 18.140 | 64.640 | 0.630 | 1.140 |
| 2080s RCP 4.5 | 3.460 | 26.540 | 18.210 | 64.770 | 0.710 | 1.270 |
| 2080s RCP 8.5 | 6.110 | 29.190 | 18.490 | 65.280 | 0.990 | 1.780 |

### June Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_C | Proj_Air_C | Obs_Baseline_Tw_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F | Obs_Max_Tw_C | Proj_Max_Tw_C | Proj_Max_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | June | 0.000 | 19.940 | 19.940 | 13.160 | 13.160 | 55.700 | 0.000 | 0.000 | 19.800 | 19.800 | 67.640 |
| 2040s RCP 4.5 | June | 2.360 | 19.940 | 22.300 | 13.160 | 14.110 | 57.400 | 0.950 | 1.710 | 19.800 | 20.750 | 69.350 |
| 2040s RCP 8.5 | June | 2.970 | 19.940 | 22.910 | 13.160 | 14.300 | 57.740 | 1.140 | 2.050 | 19.800 | 20.940 | 69.690 |
| 2080s RCP 4.5 | June | 3.460 | 19.940 | 23.400 | 13.160 | 14.440 | 57.990 | 1.270 | 2.290 | 19.800 | 21.070 | 69.930 |
| 2080s RCP 8.5 | June | 6.110 | 19.940 | 26.050 | 13.160 | 14.980 | 58.970 | 1.820 | 3.280 | 19.800 | 21.620 | 70.920 |

### July Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_C | Proj_Air_C | Obs_Baseline_Tw_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F | Obs_Max_Tw_C | Proj_Max_Tw_C | Proj_Max_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | July | 0.000 | 24.970 | 24.970 | 16.610 | 16.610 | 61.890 | 0.000 | 0.000 | 23.000 | 23.000 | 73.400 |
| 2040s RCP 4.5 | July | 2.360 | 24.970 | 27.330 | 16.610 | 16.970 | 62.540 | 0.360 | 0.650 | 23.000 | 23.360 | 74.050 |
| 2040s RCP 8.5 | July | 2.970 | 24.970 | 27.940 | 16.610 | 17.040 | 62.670 | 0.430 | 0.770 | 23.000 | 23.430 | 74.170 |
| 2080s RCP 4.5 | July | 3.460 | 24.970 | 28.430 | 16.610 | 17.080 | 62.750 | 0.480 | 0.860 | 23.000 | 23.480 | 74.260 |
| 2080s RCP 8.5 | July | 6.110 | 24.970 | 31.080 | 16.610 | 17.270 | 63.080 | 0.660 | 1.190 | 23.000 | 23.660 | 74.590 |

### August Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_C | Proj_Air_C | Obs_Baseline_Tw_C | Proj_Tw_C | Proj_Tw_F | Delta_Tw_C | Delta_Tw_F | Obs_Max_Tw_C | Proj_Max_Tw_C | Proj_Max_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | August | 0.000 | 24.350 | 24.350 | 19.620 | 19.620 | 67.310 | 0.000 | 0.000 | 25.000 | 25.000 | 77.000 |
| 2040s RCP 4.5 | August | 2.360 | 24.350 | 26.710 | 19.620 | 20.030 | 68.050 | 0.410 | 0.740 | 25.000 | 25.410 | 77.740 |
| 2040s RCP 8.5 | August | 2.970 | 24.350 | 27.320 | 19.620 | 20.110 | 68.190 | 0.490 | 0.880 | 25.000 | 25.490 | 77.880 |
| 2080s RCP 4.5 | August | 3.460 | 24.350 | 27.810 | 19.620 | 20.160 | 68.290 | 0.540 | 0.980 | 25.000 | 25.540 | 77.980 |
| 2080s RCP 8.5 | August | 6.110 | 24.350 | 30.460 | 19.620 | 20.370 | 68.670 | 0.760 | 1.360 | 25.000 | 25.760 | 78.360 |

### June Warm-Day (90th pct) Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_P90_C | Proj_Air_P90_C | Obs_Baseline_Tw_P90_C | Proj_Tw_P90_C | Proj_Tw_P90_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | June | 0.000 | 24.900 | 24.900 | 15.000 | 15.000 | 59.000 | 0.000 | 0.000 |
| 2040s RCP 4.5 | June | 2.360 | 24.900 | 27.260 | 15.000 | 15.370 | 59.660 | 0.370 | 0.660 |
| 2040s RCP 8.5 | June | 2.970 | 24.900 | 27.870 | 15.000 | 15.440 | 59.780 | 0.440 | 0.780 |
| 2080s RCP 4.5 | June | 3.460 | 24.900 | 28.360 | 15.000 | 15.480 | 59.870 | 0.480 | 0.870 |
| 2080s RCP 8.5 | June | 6.110 | 24.900 | 31.010 | 15.000 | 15.670 | 60.210 | 0.670 | 1.210 |

### July Warm-Day (90th pct) Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_P90_C | Proj_Air_P90_C | Obs_Baseline_Tw_P90_C | Proj_Tw_P90_C | Proj_Tw_P90_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | July | 0.000 | 29.100 | 29.100 | 18.900 | 18.900 | 66.020 | 0.000 | 0.000 |
| 2040s RCP 4.5 | July | 2.360 | 29.100 | 31.460 | 18.900 | 19.050 | 66.280 | 0.150 | 0.260 |
| 2040s RCP 8.5 | July | 2.970 | 29.100 | 32.070 | 18.900 | 19.070 | 66.330 | 0.170 | 0.310 |
| 2080s RCP 4.5 | July | 3.460 | 29.100 | 32.560 | 18.900 | 19.090 | 66.360 | 0.190 | 0.340 |
| 2080s RCP 8.5 | July | 6.110 | 29.100 | 35.210 | 18.900 | 19.160 | 66.490 | 0.260 | 0.470 |

### August Warm-Day (90th pct) Projections

| Scenario | Month | Delta_Air_C | Baseline_Air_P90_C | Proj_Air_P90_C | Obs_Baseline_Tw_P90_C | Proj_Tw_P90_C | Proj_Tw_P90_F | Delta_Tw_C | Delta_Tw_F |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Baseline (1997–2026) | August | 0.000 | 28.360 | 28.360 | 21.300 | 21.300 | 70.340 | 0.000 | 0.000 |
| 2040s RCP 4.5 | August | 2.360 | 28.360 | 30.720 | 21.300 | 21.470 | 70.650 | 0.170 | 0.310 |
| 2040s RCP 8.5 | August | 2.970 | 28.360 | 31.330 | 21.300 | 21.500 | 70.710 | 0.200 | 0.370 |
| 2080s RCP 4.5 | August | 3.460 | 28.360 | 31.820 | 21.300 | 21.530 | 70.750 | 0.230 | 0.410 |
| 2080s RCP 8.5 | August | 6.110 | 28.360 | 34.470 | 21.300 | 21.610 | 70.900 | 0.310 | 0.560 |

### Weekly Water Temperature Stats

| week | mean_tw | min_tw | max_tw | n_years | primary_month |
| --- | --- | --- | --- | --- | --- |
| 22.000 | 11.885 | 9.850 | 13.575 | 24.000 | 6.000 |
| 23.000 | 12.257 | 10.662 | 14.256 | 29.000 | 6.000 |
| 24.000 | 12.951 | 11.428 | 14.651 | 29.000 | 6.000 |
| 25.000 | 13.698 | 11.917 | 16.770 | 29.000 | 6.000 |
| 26.000 | 14.363 | 12.154 | 17.446 | 44.000 | 6.000 |
| 27.000 | 15.218 | 13.106 | 17.714 | 39.000 | 7.000 |
| 28.000 | 16.088 | 14.103 | 19.482 | 29.000 | 7.000 |
| 29.000 | 17.016 | 15.003 | 20.008 | 29.000 | 7.000 |
| 30.000 | 17.737 | 15.742 | 20.695 | 32.000 | 7.000 |
| 31.000 | 18.510 | 15.600 | 20.913 | 50.000 | 8.000 |
| 32.000 | 19.337 | 16.888 | 21.399 | 28.000 | 8.000 |
| 33.000 | 19.761 | 17.860 | 21.957 | 29.000 | 8.000 |
| 34.000 | 20.047 | 18.466 | 22.768 | 29.000 | 8.000 |
| 35.000 | 20.207 | 18.451 | 21.654 | 29.000 | 8.000 |
| 36.000 | 20.553 | 19.075 | 21.835 | 5.000 | 8.000 |

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
