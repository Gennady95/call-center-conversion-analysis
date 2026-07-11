# Подсчёт и контроль показателей KPI операторов колл-центра
Скрипт для контроля и подсчёта показателей KPI для операторов колл-центра. Предоставляет аналитику и инструменты для контроля за корректностью заполнения карточек клиентов и лидов в Bitrix24

# Call Center Conversion Analysis

> Python analytics tool for measuring call center conversion, monitoring CRM data quality and generating management reports.

## Description

This project analyzes the performance of a call center using Bitrix24 CRM data.

The application calculates lead conversion by operator and department, separates calls from messaging channels, detects CRM data quality issues and automatically generates Excel reports with charts.

Important:

The project requires access to a Bitrix24 CRM instance.

Without CRM access the application cannot be executed.

## Business Goal

The project helps evaluate:

- operator performance
- department conversion
- call vs message conversion
- CRM data quality
- lead processing efficiency

## Features

- Bitrix24 API integration
- Lead extraction
- Call center KPI calculation
- Conversion analysis
- Department comparison
- Operator comparison
- CRM validation
- Automatic error detection
- Excel report generation
- Plotly visualizations
- Dash dashboard components
- Telegram notifications

## Tech Stack

- Python
- pandas
- NumPy
- Plotly
- Dash
- SQLAlchemy
- fast-bitrix24
- requests
- xlsxwriter
- python-dotenv
- pyTelegramBotAPI

## How It Works

1. Downloads leads from Bitrix24
2. Extracts operators, departments and lead sources
3. Separates calls and messaging channels
4. Calculates conversion rates
5. Detects logical CRM errors
6. Creates charts
7. Generates formatted Excel reports

## Example / Demo

### Input

Bitrix24 CRM

- Leads
- Operators
- Departments
- Lead sources

### Output

Excel workbook containing:

- Lead conversion
- Call conversion
- Message conversion
- Active leads
- CRM validation errors
- KPI reports

Charts:

- Conversion by operator
- Conversion by department
- CRM error distribution

## Use Case

This project can be used for:

- call center analytics
- CRM analytics
- KPI reporting
- lead management
- business intelligence
