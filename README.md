# VBA Data Workflow Automation: Batch Reporting Pipeline

## 📌 Overview
This repository contains a VBA automation script designed to streamline enterprise-scale data reporting. It processes master datasets, dynamically extracts unique commercial entities, sanitizes data strings, and programmatically generates filtered PDF reports.

## 🛠️ Key Technical Features
- **Algorithm Efficiency:** Utilizes `Scripting.Dictionary` for fast, memory-efficient extraction of unique data points (e.g., Commercial Names) across large datasets.
- **Data Sanitization:** Programmatically cleanses target strings by utilizing `Replace` functions to remove illegal system characters, preventing runtime errors during file generation.
- **Memory Optimization:** Implements `Application.ScreenUpdating = False` and `DisplayAlerts = False` to minimize CPU and memory overhead during batch loop processing.
- **Dynamic Filtering & Export:** Programmatically loops through dynamically filtered ranges and executes batch PDF exports native to the Excel environment.

## 🚀 Business Impact
- Replaces repetitive manual data segmentation and exporting tasks.
- Eliminates human error in data matching and file nomenclature.
- Scalable logic that saves approximately 10+ hours per week in cross-regional operational reporting workflows.

*Note: All proprietary company data and specific business variables have been omitted or anonymized for compliance purposes.*
