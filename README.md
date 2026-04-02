# Sync2D_Postprocessor.py

## Description

`Sync2D_Postprocessor.py` is a post-processing pipeline for GC×GC-MS datasets exported from Sync2D workflows. It performs multi-stage data cleaning, transformation, feature consolidation, prioritization, and formatting.

Core capabilities include:

- Spectrum format normalization
- Base mass extraction and validation
- Quant/Base mass intensity ratio calculation
- Interactive unknown feature renaming
- ML-based peak area cutoff estimation to separate real chromatographic peaks from pseudo/noise peaks
- Feature merging based on RT proximity and spectral similarity
- Detection frequency (DF) and detection number (DN) calculations
- Feature prioritization based on blanks and DF thresholds
- Peak area filtering and LOD handling
- Retention index (RI) enrichment and delta calculation
- Automated Excel export with extensive formatting and visualization

The script is highly interactive and depends heavily on user input and predefined dataset structure.

**NB!** The script is tightly coupled to specific column naming conventions and dataset layouts.

## What does the script do

The script does the following:

1. Input file loading:

- Opens a GUI file selection dialog (`tkinter`)
- Accepts `.xlsx`, `.csv`, or `.txt` files
- Loads into a `pandas DataFrame (df)`
- Prints `df.info()` and First 5 rows
- Hard stop if no file is selected.

2. Base mass retrieval (optional). If enabled:

- Requires: `"Quant mass"`, `"Spectrum"` or `"Spectrum_Sync2D"`
- Extracts base mass: Parses spectrum (two formats supported: `(mz|intensity)` or `mz:intensity`)
- Identifies highest intensity fragment → Base mass
- Creates `"Base mass"` column if missing
- Validates `"Quant mass"`: Replaces invalid (`NaN`, `0`, non-numeric) with `Base mass`
- Ensures `Quant mass` exists in spectrum
- Computes `"Int(QM):Int(BM) (%)"`

3. Unknown feature renaming (optional). If enabled:

- User defines: Keyword identifying unknowns (`Feature`, `Peak`, etc.) and two identifier columns (e.g., `"R.I. calc"`, `"Med RT2 (sec)"`)
- Renames: `Feature → Feature_<ID1>_<ID2>`.
- ID formatting: `ID1 → integer`, `ID2 → 3 decimal float`.

4. Peak Area Cutoff Estimation (optional, exploratory QC step). The script includes an advanced statistical module for estimating a cutoff threshold that separates real chromatographic peaks from pseudo/noise peaks based on peak area distributions. This step is interactive and must be explicitly enabled by the user.

- User prompt: `Would you like to perform real vs. pseudo peak cutoff estimation? 0 = No 1 = Yes`. If `0` → section is skipped entirely; if `1` → full cutoff estimation workflow is executed; `Invalid input` → skipped with warning.
- Uses all detected area columns (`area_columns`)
- Flattens values into a single vector `area_values = df[area_columns].values.flatten()`
- Cleans the data: Removes NaN values and removes zeroes (assumed to represent absence of peaks)
- Computes distribution summary: Mean, standard deviation, quartiles (25%, 50%, 75%), high percentiles (90%, 95%, 99%)
- Printed to console for inspection: `pd.Series(area_values).describe(...)`
- Histogram visualization (log10 scale): (i) Applies log10 transformation `log_area_values = np.log10(area_values)` and (ii) plots histogram (density-normalized) to reveal bimodal structure (`pseudo vs. real peaks`) and identify approximate separation region.
- Kernel Density Estimation (KDE): (i) Performs non-parametric density estimation with `gaussian_kde(log_area, bw_method=0.2)`; (ii) detects local minima in KDE curve that represents valley between low-intensity (noise/pseudo peaks) and high-intensity (true peaks); (iii) If minimum exists, defines `cutoff_kde_log10` and `cutoff_kde` (converted back to linear scale); (iv) If no minimum found, KDE cutoff is set to `None` and warning is printed.
- Gaussian Mixture Model (GMM): (i) Fits a 2-component Gaussian Mixture Model with `GaussianMixture(n_components=2)`; (ii) Assumes: Component 1 → pseudo peaks (low mean); Component 2 → real peaks (high mean); (iii) Extracts: Means (μ1, μ2), Standard deviations (σ1, σ2); (iv) Computes analytical intersection point between the two Gaussians; This represents the GMM-based cutoff; (v) Outputs: `cutoff_gmm_log10` and `cutoff_gmm`.
- Combined cutoff definition. The script defines three cutoff values: (i) `cutoff_high` (recommended, conservative): `max(GMM, KDE)`; (ii) cutoff_low: `min(GMM, KDE)`; (iii) cutoff_mean: Average of GMM and KDE cutoffs; special case: if KDE fails, all three cutoffs default to `cutoff_gmm`.
- Visualization of results. A combined diagnostic plot is generated showing: Histogram (log10 peak areas), KDE curve, two GMM component distributions (pseudo peaks and real peaks), vertical lines for GMM cutoff, KDE cutoff (if available), and mean cutoff. This visualization provides a clear analytical basis for selecting filtering thresholds.
- Output and usage. Cutoff values are printed to console, but not automatically enforced downstream. Intended usage: manual inspection, user-defined filtering in later processing steps, and parameter tuning for feature selection workflows.

5. Spectrum conversion:

- Converts Sync2D spectrum format `(77|871.48)(156|1000.00)` to `77:871.48 156:1000.00`
- Renames: `"Spectrum"` → `"Spectrum_Sync2D"`
- Inserts new `"Spectrum"` column. User specifies insertion position.

6. Feature merging (unknowns). If enabled:

- Targets features matching unknown keyword
- Groups features based on `RT1 tolerance` (±10 sec default), `RT2 tolerance` (±0.5 sec default), optional `Base mass equality`, `spectral similarity` using either `"DISCO"` (mean-centered cosine) or `NDP"` (dot product) with a default similarity threshold of `0.7`
- Merging logic: Identify groups → Keep feature with highest total area → Sum all sample areas into main feature → Drop others → Track merged IDs
- Outputs: Merged dataset and optional merged/unmerged export.

7. Detection frequency (DF) & DN calculations (optional). If enabled:

- Uses predefined column groups: UMU blanks, field blanks, country-specific samples
- Calculates: `"DN"` = count of detections above cutoff, `"DF (%)"` = frequency %, per-country metrics `"DN CZ"`, `"DF CZ (%)"`, etc.

8. Feature prioritization (optional). Based on:

- Area ratios vs. blanks
- Detection frequency
- Unknown feature behavior
- Country-level DF
- Outputs: `"Priority"` (0, 1, 2), `"Priority_1–4"` (step-wise logic), `"Reason"` (traceable logic), `"Ranking"`: `2 = keep`, `1 = borderline`, `0 = delete`.

9. LOD specification & cutoff selection. If enabled:

- Uses: Precomputed cutoffs (if available) OR Manual user input
- Populates `"LOD"` column.

10. Replace small peak areas. If enabled:

- Preserves original data: Creates `*_INIT` columns
- Replaces: `Area < cutoff` OR `NA` → `-cutoff`
- Applies to all sample columns
- Rounds values to integers.

11. Retention index processing (optional). If enabled:

- Populates `"R.I. lib"` from `"RI_Semi-Std_NP"`, `"RI_Std_NP"`, `"RI_AI"`
- Adds `"R.I. source"`
- Computes `R.I. delta = R.I. calc − R.I. lib`.

12. Output generation.

- Full processed dataset: `*_Prcssd.xlsx`.

13. Applies extensive Excel formatting using `openpyxl`:

- Column coloring by type
- Conditional formatting (e.g., intensity thresholds)
- Column width adjustments
- Header formatting
- Freeze panes and zoom.

## Prerequisites

Before using the script, several applications/tools have to be installed:

1. Visual Studio Code; https://code.visualstudio.com/download.
2. Python 3; https://www.python.org/downloads/windows/.
3. Python Extension in Visual Studio Code > Extensions (`Ctrl + Shift + X`) > Search “python” > Press `Install`.

Then, the required packages, i.e. `pandas`, `numpy`, `matplotlib`, `scikit-learn`, `scipy`, and `openpyxl`, must be installed as follows:
Visual Studio Code > Terminal > New Terminal > In terminal, type `pip install pandas numpy matplotlib openpyxl scikit-learn scipy` > Press `Enter`.

## How to use the script

To use the script, the following steps must be executed:

1. Run the script:

- Right mouse click anywhere in Visual Studio Code script file > Run Python > Run Python File in Terminal or press `play` button in the top-right corner.

2. Select input files:

- A file dialog window will appear
- Select main dataset (`CSV`, `TXT`, or `Excel` file)
- Click `Open`.

3. Respond to prompts:

- Enable/disable processing steps
- Provide column names (if needed)
- Provide cutoff values (if needed).

## Notes and recommendations

**Input data requirements**

1. Main dataset - Critical columns:

- `"Name"`
- `"Quant mass"`
- Multiple sample area columns
- Strongly required: `"Spectrum"` or `"Spectrum_Sync2D"` ,`"Med RT1 (sec)"`, `"Med RT2 (sec)"` (for merging), `"Area ave."`, 
- Required for advanced features: `"Base mass"` (optional, otherwise computed), `"Keep"` (for prioritization exclusions), `"CL Sync2D"` (priority logic), `"TargetHit"` (priority logic), `"R.I. calc"`
- Optional but recommended: `"RI_Semi-Std_NP"`, `"RI_Std_NP"`, `"RI_AI"`, `"CAS"`.

2. Sample area columns

- Multiple columns representing samples
- Used for merging, DF/DN calculations, area filtering
- Must be numeric or convertible.

3. Blank columns (hardcoded)- Examples:

- "Blank_01_241112_012"
- "CZ_FB_01_241112_017"

These are used in DF calculations and prioritization.

**Key processing dependencies**

Base mass retrieval: `Quant mass, Spectrum`
Spectrum conversion: `Spectrum`
Feature merging: `Spectrum, RT1, RT2`
DF calculations: `Sample + blank columns`
Prioritization: `Keep, DF columns`
RI processing: `R.I. calc + RI sources`

**Limitations**

- No fuzzy matching.
- Hardcoded column names (especially blanks and countries).
- Requires consistent naming conventions.
- Interactive prompts reduce automation.
- Computationally expensive for large datasets (spectral comparisons).
- Partial dependency on user-provided column names.
- Cutoff estimation assumes bimodal peak area distribution; may fail for unimodal or highly skewed datasets.
- KDE minimum detection is sensitive to bandwidth (bw_method) and data density.
- GMM assumes Gaussian-like distributions and may misrepresent heavy-tailed data.
- Computation may be slow for very large datasets due to KDE and GMM fitting.
- Results are not automatically applied — user must manually integrate cutoff into filtering logic.

## License

[![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/license/mit)

Intended for academic and research use.
