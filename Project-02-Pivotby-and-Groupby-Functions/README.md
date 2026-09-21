# Project 02 – PIVOTBY & GROUPBY Functions

## Overview

This project explores Excel's `GROUPBY()` and `PIVOTBY()` functions to create dynamic, formula-based reports from sales data.

The project demonstrates how modern Excel dynamic-array functions can be used to summarize data, calculate percentages, create cross-tabulated reports, and build reports that automatically update when the underlying data changes.

## Objectives

* Summarize sales data using `GROUPBY()`
* Create cross-tabulated reports using `PIVOTBY()`
* Calculate total sales by category
* Analyze sales by payment method
* Calculate percentages of row, column, and grand totals
* Build dynamic reports controlled by worksheet inputs
* Understand the differences between PivotTables and formula-based reporting

## Tools & Functions

* Microsoft Excel
* `GROUPBY()`
* `PIVOTBY()`
* `SUM`
* `PERCENTOF`
* `HSTACK()`
* `VSTACK()`
* Dynamic spilled-array formulas
* Conditional formatting
* Data validation

## Analysis

### 1. GROUPBY Sales Analysis

`GROUPBY()` was used to summarize sales data based on row-level categories.

The function groups records and applies an aggregation function such as `SUM` to calculate the resulting values.

This creates a dynamic report without requiring a traditional PivotTable.

### 2. PIVOTBY Cross-Tab Analysis

`PIVOTBY()` was used to create a cross-tabulated report with variables placed in both the row and column fields.

This allows sales performance to be analyzed across multiple dimensions simultaneously.

### 3. Percentage Analysis

`PERCENTOF` was used to calculate sales percentages relative to different totals.

The analysis includes:

* Percentage of Column Total
* Percentage of Row Total
* Percentage of Grand Total

### 4. Dynamic Report Controls

A worksheet input was connected to the `PIVOTBY()` calculation so that the user can change the percentage calculation without manually rebuilding the report.

This demonstrates how Excel formulas can be used to create interactive reporting tools.

## Key Excel Concepts

### Dynamic Array Formulas

The reports are generated from formulas stored in a single cell and automatically spill into the required range.

This allows the report to expand or contract based on the underlying data.

### GROUPBY vs. PIVOTBY

`GROUPBY()` is useful when the analysis primarily requires grouping records by ro
