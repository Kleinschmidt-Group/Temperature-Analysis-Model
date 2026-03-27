Chief Joseph Dam Tailrace Water Temperature Analysis
Historical Trend Analysis & Future Climate Projections
Kleinschmidt Associates | 2026


PURPOSE
-------
This script performs a complete statistical analysis of July–August water
temperatures at the Chief Joseph Dam (CJD) tailrace on the Columbia River.
It answers two questions:

  1. Have summer water temperatures changed over the historical record?
  2. How much could they change in the future under climate warming scenarios?

The analysis produces a formatted Excel workbook of results and 15
publication-quality figures.


REQUIREMENTS
------------
Python 3.8+

Required packages:
  numpy
  pandas
  matplotlib
  scipy
  openpyxl

Install with:
  pip install numpy pandas matplotlib scipy openpyxl


INPUT DATA
----------
The script expects five required input files plus two optional MACA climate
files. File paths are set in the CONFIGURATION section at the top of
CJD_Analysis_Test.py (lines ~139–173). Update these paths if data is moved.

Water temperature data (4 Excel files, one per station):
  FDRW_Hourly_Data_1995_2025.xlsx  — Grand Coulee Dam Forebay
  GCGW_Hourly_Data_1995_2025.xlsx  — Grand Coulee Dam Tailrace
  CHJ_Hourly_Data_1995_2025.xlsx   — Chief Joseph Dam Forebay
  CHQW_Hourly_Data_1995_2025.xlsx  — Chief Joseph Dam Tailrace (primary)

  Source:  DART (Data Access in Real Time)
           Columbia Basin Research, University of Washington
           http://www.cbr.washington.edu/dart/query/river_graph_wmq
  Format:  Multi-sheet Excel workbook (one sheet per year)
           Columns: Pacific Timestamp, Temperature (C), Temperature (F),
                    Outflow (kcfs)
  Period:  CHQW: 1997–2025 (after QA/QC removes 1995–1996 null values)
           FDRW, GCGW, CHJ: 2000–2025

Air temperature data (1 CSV file):
  Douglas_Temp_1995_2025.csv

  Source:  NOAA National Centers for Environmental Information (NCEI)
           Global Historical Climatology Network Daily (GHCND)
  Station: DOUGLAS WASHINGTON, WA US (GHCND:USR0000WDOU)
           47.619 N, 119.899 W -- approximately 10 miles from CJD
  Format:  CSV with columns DATE, TMAX, TMIN, TAVG (all in deg F)
           Daily mean computed as (TMAX + TMIN) / 2 after conversion to deg C
  Period:  1995-2025 subset used; full station record through 2026-02-21

MACAv2-METDATA climate files (2 CSV files -- optional):
  FutureTimeSeriesMax.csv  -- JJA maximum temperature, all CMIP5 models
  FutureTimeSeriesMin.csv  -- JJA minimum temperature, all CMIP5 models

  Source:  Climate Toolbox (climatetoolbox.org), MACAv2-METDATA
           (Abatzoglou & Brown 2012), 20 CMIP5 GCMs
  Grid:    47.619 N, 119.899 W (Douglas County, WA)
  Note:    If these files are not present, the script uses pre-computed
           delta values embedded in the code (see CLIMATE SCENARIOS below).


HOW TO RUN
----------
1. Verify that all required input data files are in place and that the file
   paths in the CONFIGURATION section point to them.

2. Verify that OUTPUT_DIR points to the desired output folder. The script
   will create this folder if it does not exist.

3. Run from the command line:

     python CJD_Analysis_Test.py

4. The script prints progress to the console and saves all outputs to
   OUTPUT_DIR.

Expected runtime: under 5 minutes on a standard workstation.


OUTPUTS
-------
All outputs are saved to the folder specified by OUTPUT_DIR.

Excel workbook:
  CJD_Temperature_Results.xlsx
    Sheet 1: Data Sources & Methods     -- full documentation of inputs and methods
    Sheet 2: Trend Summary (Jul-Aug)    -- combined July-August Mann-Kendall and
                                           Theil-Sen results for all stations
    Sheet 3: Trend Summary (Monthly)    -- same results split by July and August
                                           separately for all stations
    Sheet 4: Monthly Projections        -- projected water temperatures for July
                                           and August separately under each
                                           climate scenario
    Sheet 5: Projections (Jul-Aug)      -- combined Jul-Aug projections (retained
                                           for reference)
    Sheet 6: CHQW Annual                -- year-by-year July mean, August mean,
                                           and combined July-August mean water
                                           temperatures at CHQW
    Sheet 7: Mohseni Parameters         -- fitted model parameters (mu, alpha,
                                           gamma, beta) and NSE for all-weeks
                                           and Jul-Aug-only fits
    Sheet 8: Weekly Stats               -- observed and projected week-by-week
                                           mean, min, and max temperatures at
                                           CHQW for July and August

Figures (PNG, 150 dpi):
  Fig1_Annual_Summer_Trends.png         -- annual July and August mean water
                                           temperature at all four stations,
                                           split by month with Theil-Sen trends
                                           (4-row x 2-column panel layout)
  Fig2_Seasonal_Climatology.png         -- mean annual temperature cycle (7-day
                                           rolling average) for all stations and air
  Fig3_CHQW_Trend_Detail.png            -- 5-panel figure: July water trend,
                                           August water trend, July observed vs.
                                           Mohseni predicted, August observed vs.
                                           Mohseni predicted, and air temperature
                                           trend (both months), all with bootstrap
                                           95% confidence bands
  Fig4_JulAug_Daily_By_Year.png         -- daily July-August water temperatures
                                           at CHQW, all years overlaid and
                                           colored by year, with multi-year
                                           average daily cycle
  Fig5_Mohseni_Regression.png           -- weekly air vs. water temperature
                                           scatter colored by season, with
                                           fitted Mohseni S-curves (all weeks
                                           and Jul-Aug only)
  Fig6_Mohseni_Observed_vs_Predicted.png -- observed vs. Mohseni-predicted
                                           weekly water temperature scatter,
                                           colored by year, with NSE and R-squared
  Fig7_JulAug_Obs_vs_Pred_TimeSeries.png -- July and August annual observed vs.
                                           Mohseni-predicted means side by side,
                                           with mean absolute error (MAE)
  Fig8_Mohseni_All_Years_Overlaid.png   -- all years of weekly air-water data
                                           overlaid on the fitted Mohseni S-curve,
                                           colored by year
  Fig9_Projected_Temps_Bar.png          -- projected July and August water
                                           temperatures under each climate
                                           scenario, bar chart with deg C and
                                           deg F labels
  Fig10_Projection_Points_Mohseni.png   -- climate projection points (July and
                                           August) plotted on the Mohseni
                                           S-curve for all scenarios
  Fig11_Period_Comparison_Boxplots.png  -- monthly boxplot distributions comparing
                                           early vs. recent periods at all four
                                           stations, with Mann-Whitney p-values
                                           for July and August
  Fig12_Monthly_Projections_Bar.png     -- projected change (delta-Tw) in July
                                           and August water temperature by
                                           climate scenario, with observed
                                           baselines annotated
  Fig13_Monthly_Projection_Points.png   -- Mohseni model sensitivity diagram
                                           showing delta arrows from baseline
                                           to 2080s RCP 8.5 for July and August
  Fig14_Observed_Weekly_MinMaxMean.png  -- observed week-by-week mean, min, and
                                           max water temperatures at CHQW for
                                           July through August (ISO weeks)
  Fig15_Projected_Weekly_MinMaxMean.png -- observed vs. projected week-by-week
                                           water temperatures under the most
                                           extreme scenario (2080s RCP 8.5)


ANALYSIS OVERVIEW
-----------------
The script executes four analytical steps:

Step 1 -- QA/QC and Data Loading
  All water temperature and air temperature records are screened by the
  TemperatureQAQC class before analysis. Five checks are applied: physical
  plausibility bounds (by month), IQR-based outlier detection (3x IQR),
  constant-value spans indicative of sensor failure (>=24-hour duration),
  impossible hourly jumps (>3 deg C/hr, per EPA 2002), and data gap
  identification. Flagged records are removed prior to analysis.

Step 2 -- Annual Summer Averages and Trend Analysis
  For each station and each year, mean water temperatures are computed
  separately for July and August (and combined) across all hourly readings.

  Mann-Kendall trend test: A nonparametric test that determines whether there
  is a statistically significant upward or downward trend over time. Trend is
  considered significant if p < 0.05.
  Reference: Hirsch et al. (1991); Helsel & Hirsch (2002)

  Theil-Sen slope estimator: Calculates the median pairwise slope across all
  year combinations, resistant to outliers. Returns a slope in deg C/year with
  a 95% confidence interval.
  Reference: Helsel & Hirsch (2002)

  Trends are computed for all four stations, both for combined July-August
  and for each month separately.

Step 3 -- Mohseni Air-Water Temperature Model
  Fits the Mohseni et al. (1998) nonlinear logistic regression to weekly
  air-water temperature pairs at CHQW:

    Tw = mu + (alpha - mu) / (1 + exp(gamma * (beta - Ta)))

  Parameters:
    mu (minimum)   -- minimum water temperature (deg C), the winter floor
    alpha (maximum) -- maximum water temperature (deg C), the thermal ceiling
    gamma           -- steepness of the S-curve
    beta            -- air temperature at the inflection point (deg C)

  Two fits are produced: one using all weekly data and one restricted to
  July-August weeks (ISO weeks 26-35). Model fit is evaluated using
  Nash-Sutcliffe Efficiency (NSE; Nash & Sutcliffe 1970).
  Reference: Mohseni et al. (1998); Mantua et al. (2010)

Step 4 -- Future Climate Projections
  Applies CMIP5 air temperature warming deltas to the fitted Mohseni model
  via the delta method. July and August are projected separately using their
  respective observed baseline air temperatures.

  Climate scenarios (MACAv2-METDATA, 20-model CMIP5 ensemble,
  Abatzoglou & Brown 2012, accessed via Climate Toolbox):
    2040s RCP 4.5 (moderate emissions): +2.36 deg C air warming
    2040s RCP 8.5 (high emissions):     +2.97 deg C air warming
    2080s RCP 4.5 (moderate emissions): +3.46 deg C air warming
    2080s RCP 8.5 (high emissions):     +6.11 deg C air warming

  Deltas are referenced to the 1970-1999 model historical baseline period.
  If the MACAv2 CSV files are present at the paths defined in CONFIGURATION,
  deltas are recomputed at runtime; otherwise the pre-computed values above
  are used automatically.

  NOTE: These CMIP5/RCP scenarios replace the CMIP3-era A1B/B1 scenarios
  from Mantua et al. (2010) used in earlier versions of this script.

  Reference: Abatzoglou & Brown (2012); Mantua et al. (2010)


CODE STRUCTURE
--------------
Line     Section
------   ------------------------------------------
1-111    Header documentation and references
113-129  Imports
131-269  Configuration (file paths, MACA delta computation, station labels,
           climate scenarios, color palette)
271-292  Matplotlib style settings
294-371  Section 1: Statistical helper functions
           mann_kendall()  -- trend test
           theil_sen()     -- slope estimator
           mohseni()       -- air-water model equation
           fit_mohseni()   -- curve fitting + NSE
374-~500 Section 1B: QA/QC (TemperatureQAQC class)
           check_physical_bounds()    -- monthly plausibility bounds
           detect_outliers_iqr()      -- IQR-based outlier flagging
           detect_constant_spans()    -- sensor failure detection
           detect_impossible_jumps()  -- hourly jump screening
           detect_data_gaps()         -- gap identification
~500+    Section 2: Data loading
           load_station()  -- reads multi-sheet water temp Excel files,
                              applies QA/QC, returns cleaned DataFrame
           load_air()      -- reads NOAA daily air temp CSV, applies QA/QC,
                              converts deg F to deg C
~650+    Section 3: Analysis
           run_analysis()  -- computes annual means by month and combined,
                              fits Mohseni model (all weeks and Jul-Aug only),
                              runs trend tests for all stations and months,
                              generates monthly projections and week-by-week
                              temperature statistics
~920+    Section 4: Excel export
           export_excel()  -- creates formatted 8-sheet workbook
~1164+   Section 5: Figures
           make_figures()  -- generates 15 publication-quality PNG figures
~2099+   Section 6: Main entry point


CONFIGURATION
-------------
To modify the analysis, edit the CONFIGURATION section (~lines 131-269):

  OUTPUT_DIR        -- where results are saved
  STATION_FILES     -- paths to the four water temperature Excel files
  AIR_TEMP_FILE     -- path to the NOAA air temperature CSV
  MACA_TMAX_FILE    -- path to Climate Toolbox JJA Tmax CSV (optional)
  MACA_TMIN_FILE    -- path to Climate Toolbox JJA Tmin CSV (optional)
  SUMMER_MONTHS     -- months included in the summer average (default: [7, 8])
  STATION_ORDER     -- order stations appear in figures
  CLIMATE_SCENARIOS -- warming increments and scenario labels; auto-populated
                       from MACA files if present, otherwise pre-computed


REFERENCES
----------
Abatzoglou, J.T. and Brown, T.J. (2012). A comparison of statistical
  downscaling methods suited for wildfire applications. International Journal
  of Climatology, 32, 772-780.  [MACAv2-METDATA downscaling method]

Helsel, D.R. and Hirsch, R.M. (2002). Statistical Methods in Water Resources.
  USGS Techniques of Water-Resources Investigations, Book 4, Chapter A3.

Hirsch, R.M., Alexander, R.B., and Smith, R.A. (1991). Selection of methods
  for the detection and estimation of trends in water quality. Water Resources
  Research, 27(5), 803-813.

Mantua, N., Tohver, I., and Hamlet, A. (2010). Climate change impacts on
  streamflow extremes and summertime stream temperature and their possible
  consequences for freshwater salmon habitat in Washington State. Climatic
  Change, 102, 187-223.

Mohseni, O., Stefan, H.G., and Erickson, T.R. (1998). A nonlinear regression
  model for weekly stream temperatures. Water Resources Research, 34(10),
  2685-2692.

Nash, J.E. and Sutcliffe, J.V. (1970). River flow forecasting through
  conceptual models, Part I: A discussion of principles. Journal of Hydrology,
  10(3), 282-290.

O'Connor, J.E. et al. (2021). Water temperature trends in the Columbia River
  Basin. USACE, Institute for Water Resources.

Vermeyen, T.B. (2000). Using selective withdrawal to control reservoir
  releases and downstream temperatures. Bureau of Reclamation, PAP-0854.

Washington Climate Change Impacts Assessment (WACCIA), Chapter 6 -- Salmon.
  Mantua et al. (contributing authors).

================================================================================

