# Health Canada DIN Lookup Automation

## Overview

This project is a Google Sheets automation tool developed for a medical company to retrieve Drug Identification Numbers (DINs) from the Health Canada Drug Product Database API. The tool uses drug name, strength, and form to accurately match and populate DINs into the company's database, helping standardize their medication data.

## Business Impact

* **Saved ~10 man-hours per week** previously spent on manual DIN lookups.
* Added DIN numbers to an extensive database of **6,000+ medications**.
* Reduced data errors by **up to 35%**.
* Created a **standardized database** that became the main source of truth for multiple other systems, improving overall data accuracy and consistency.

## Demo

A short GIF demonstration of the tool in action is available below:

![DIN Lookup Demo](demo.gif)

> Note: Full 1.5-minute video demo is available upon request.

## Features

* Fetches and caches Health Canada drug data to handle large datasets efficiently.
* Supports both brand and generic medications, including combination products.
* Implements fuzzy matching and strength normalization for accurate DIN assignment.
* Handles batch processing with time-based triggers to prevent execution timeouts.
* Automatically updates Google Sheets with matched DINs and NDC codes.
* Includes error handling and logging for easy monitoring.

## Technology Stack

* **Google Apps Script** for automation within Google Sheets.
* **Health Canada APIs** for drug, status, and active ingredient data.
* **Caching** and batch processing using `CacheService` and triggers.
* JSON parsing and data normalization for matching.

## Usage

1. Open the Google Sheet containing your medication data.
2. Go to `Extensions -> Apps Script`.
3. Paste the code from `Code.gs` or link your GitHub version.
4. Run `startNdcPopulation()` to begin processing.
5. The script will populate DINs and NDCs, automatically resuming if it hits execution limits.

## Repository Structure

* `Code.gs` — Main automation script.
* `LICENSE` — MIT License.
* `README.md` — Project description and documentation.

## Notes

* Designed for large datasets and extensive medication records.
* Scalable to support integration with other downstream databases.

---

This repository showcases practical experience in healthcare data automation, API integration, and workflow optimization — demonstrating the ability to solve real-world business problems with code.
