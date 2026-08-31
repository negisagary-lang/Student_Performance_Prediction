"""
=====================================================================
EDA Framework - Reusable Exploratory Data Analysis Functions
=====================================================================
Project       : Student Performance Analysis & Prediction
Task          : Week 2 - EDA & Visualization Framework Design
Author        : [Student Name]
Description   :
    A reusable, general-purpose Exploratory Data Analysis (EDA)
    framework built in Python. It provides functions for:
      - Dataset inspection and summary
      - Missing value analysis
      - Duplicate detection
      - Statistical summarization
      - Outlier detection
      - Categorical / numerical summaries
      - Correlation analysis

    The framework is designed to adapt to ANY structured dataset,
    not just the Student Performance dataset.
=====================================================================

No real dataset is provided for this task. The functions below are
generic and will be applied once a dataset is supplied in a later phase.

Usage Example:
    import eda_framework as eda

    df = pd.read_csv("students.csv")
    eda.dataset_summary(df)
    eda.missing_value_report(df)
    eda.detect_outliers(df)
"""

import numpy as np
import pandas as pd
import warnings

# Suppress non-critical warnings for cleaner output
warnings.filterwarnings("ignore")

# ====================================================================
# 1. DATASET SUMMARY
# ====================================================================
def dataset_summary(df: pd.DataFrame) -> None:
    """
    Provide an overall summary of the dataset:
      - Shape (rows, columns)
      - Column names
      - Data types
      - Number of missing values per column
      - Basic count of unique values per column

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.
    """
    print("=" * 60)
    print("DATASET SUMMARY")
    print("=" * 60)
    print(f"Number of rows    : {df.shape[0]}")
    print(f"Number of columns : {df.shape[1]}")
    print("\nColumn names:", list(df.columns))
    print("\nData types:\n")
    print(df.dtypes)
    print("\nUnique values per column:\n")
    print(df.nunique())
    print("\nMemory usage:")
    print(df.memory_usage(deep=True))


# ====================================================================
# 2. MISSING VALUE REPORT
# ====================================================================
def missing_value_report(df: pd.DataFrame) -> pd.DataFrame:
    """
    Produce a report of missing values for every column.

    Returns a DataFrame with:
      - Count of missing values
      - Percentage of missing values
      - Data type of each column

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.

    Returns
    -------
    pd.DataFrame
        A missing-value summary report.
    """
    missing_count = df.isnull().sum()
    missing_percent = (df.isnull().sum() / len(df)) * 100
    report = pd.DataFrame(
        {
            "Missing_Count": missing_count,
            "Missing_Percent": missing_percent,
            "Data_Type": df.dtypes,
        }
    )
    report = report[report["Missing_Count"] > 0].sort_values(
        "Missing_Percent", ascending=False
    )

    print("=" * 60)
    print("MISSING VALUE REPORT")
    print("=" * 60)
    if report.empty:
        print("No missing values detected in the dataset.")
    else:
        print(report)
    return report


# ====================================================================
# 3. DUPLICATE DETECTION
# ====================================================================
def duplicate_report(df: pd.DataFrame, subset: list = None) -> int:
    """
    Detect duplicate rows (optionally on a subset of columns).

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.
    subset : list, optional
        Columns to consider when finding duplicates.
        If None, all columns are used.

    Returns
    -------
    int
        Number of duplicate rows.
    """
    dup_count = df.duplicated(subset=subset).sum()
    print("=" * 60)
    print("DUPLICATE REPORT")
    print("=" * 60)
    print(f"Number of duplicate rows : {dup_count}")
    if subset is not None:
        print(f"Considered columns        : {subset}")
    return dup_count


# ====================================================================
# 4. STATISTICAL SUMMARY
# ====================================================================
def numerical_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute descriptive statistics for all numerical columns.

    Statistics include: count, mean, std, min, quartiles, max.

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.

    Returns
    -------
    pd.DataFrame
        Descriptive statistics of numerical columns.
    """
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    print("=" * 60)
    print("NUMERICAL SUMMARY")
    print("=" * 60)
    print(f"Numerical columns: {numeric_cols}")
    result = df[numeric_cols].describe().T
    print(result)
    return result


def categorical_summary(df: pd.DataFrame) -> None:
    """
    Summarize all categorical columns by printing value counts
    (counts and percentages) for each column.

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.
    """
    cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    print("=" * 60)
    print("CATEGORICAL SUMMARY")
    print("=" * 60)
    print(f"Categorical columns: {cat_cols}")
    for col in cat_cols:
        print(f"\n--- {col} ---")
        counts = df[col].value_counts()
        percents = df[col].value_counts(normalize=True) * 100
        summary = pd.DataFrame({"Count": counts, "Percent": percents})
        print(summary)


# ====================================================================
# 5. OUTLIER DETECTION (IQR Method)
# ====================================================================
def detect_outliers(
    df: pd.DataFrame, columns: list = None, threshold: float = 1.5
) -> pd.DataFrame:
    """
    Detect outliers in numerical columns using the IQR (Interquartile
    Range) method.

    For a given column:
        Q1 = 25th percentile
        Q3 = 75th percentile
        IQR = Q3 - Q1
        Lower Bound = Q1 - threshold * IQR
        Upper Bound = Q3 + threshold * IQR

    Any value outside [Lower Bound, Upper Bound] is considered an outlier.

    IMPORTANT: This function only flags outliers. It does NOT delete
    them. Whether to remove, cap, or keep them depends on domain context.

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.
    columns : list, optional
        Numerical columns to check. Defaults to all numerical columns.
    threshold : float
        IQR multiplier (default 1.5).

    Returns
    -------
    pd.DataFrame
        A report of outlier counts and boundaries per column.
    """
    if columns is None:
        columns = df.select_dtypes(include=[np.number]).columns.tolist()

    print("=" * 60)
    print("OUTLIER REPORT (IQR METHOD)")
    print("=" * 60)

    report_rows = []
    for col in columns:
        s = df[col]
        q1 = s.quantile(0.25)
        q3 = s.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - threshold * iqr
        upper = q3 + threshold * iqr
        outlier_mask = (s < lower) | (s > upper)
        n_outliers = outlier_mask.sum()
        report_rows.append(
            {
                "Column": col,
                "Q1": round(q1, 2),
                "Q3": round(q3, 2),
                "IQR": round(iqr, 2),
                "Lower": round(lower, 2),
                "Upper": round(upper, 2),
                "Outlier_Count": n_outliers,
                "Outlier_Percent": round((n_outliers / len(s)) * 100, 2),
            }
        )

    report = pd.DataFrame(report_rows)
    print(report)
    return report


# ====================================================================
# 6. CORRELATION ANALYSIS
# ====================================================================
def correlation_matrix(df: pd.DataFrame, method: str = "pearson") -> pd.DataFrame:
    """
    Compute the correlation matrix of numerical columns.

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.
    method : str
        Correlation type: 'pearson' (linear) or 'spearman' (monotonic).

    Returns
    -------
    pd.DataFrame
        Correlation matrix.
    """
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    print("=" * 60)
    print(f"CORRELATION MATRIX ({method.upper()})")
    print("=" * 60)
    corr = df[numeric_cols].corr(method=method)
    print(corr)
    return corr


# ====================================================================
# 7. HIGH-LEVEL SAMPLE WORKFLOW
# ====================================================================
def run_full_eda(df: pd.DataFrame) -> None:
    """
    Run a standard EDA workflow on the provided dataset.

    This ties all the functions together in one call.

    Parameters
    ----------
    df : pd.DataFrame
        The input dataset.
    """
    print("\n" + "#" * 70)
    print("# FULL EDA WORKFLOW")
    print("#" * 70 + "\n")
    dataset_summary(df)
    missing_value_report(df)
    duplicate_report(df)
    numerical_summary(df)
    categorical_summary(df)
    detect_outliers(df)
    correlation_matrix(df)


# ====================================================================
# DEMO / SELF-TEST (Illustrative synthetic data)
# ====================================================================
if __name__ == "__main__":
    # -------------------------------
    # ILLUSTRATIVE EXAMPLE ONLY
    # -------------------------------
    # This creates synthetic data to demonstrate the framework.
    # THIS IS NOT REAL STUDENT DATA and no real findings are claimed.
    rng = np.random.default_rng(42)

    sample = pd.DataFrame(
        {
            "Student_ID": range(1, 101),
            "Age": rng.integers(14, 18, 100),
            "Gender": rng.choice(["Male", "Female"], 100),
            "Study_Hours": rng.normal(5, 2, 100).clip(0, 15),
            "Attendance": rng.normal(85, 10, 100).clip(0, 100),
            "Previous_Score": rng.normal(70, 15, 100).clip(0, 100),
            "Internet_Access": rng.choice(["Yes", "No"], 100),
            "Final_Score": rng.normal(75, 12, 100).clip(0, 100),
        }
    )

    # Add a deliberate missing value to demonstrate the missing report
    sample.loc[5, "Study_Hours"] = np.nan

    run_full_eda(sample)
