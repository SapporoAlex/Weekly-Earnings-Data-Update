# Kyoto Machiya Earnings Report Automation

This repository contains a script developed to automate the generation of weekly earnings reports for Kyoto Machiyas. The automation significantly improved the efficiency and accuracy of sales data reporting for the company.

<img src="https://github.com/SapporoAlex/Weekly-Earnings-Data-Update/blob/main/preview.jpg" align="center" margin-top="10" margin-bottom="10">
<figcaption><a href="https://github.com/SapporoAlex/Weekly-Earnings-Data-Update/blob/main/preview%20larger.jpg">larger</a></figcaption>

## Table of Contents

- [Project Overview](#project-overview)
- [How It Works](#how-it-works)
- [Benefits to the Company](#benefits-to-the-company)
- [Future Enhancements](#future-enhancements)
- [Acknowledgements](#acknowledgements)

## Project Overview

The script in this repository was created to address the need for timely and accurate weekly earnings reports for Kyoto Machiyas. Manually compiling these reports was time-consuming and prone to errors. By automating the process, the company was able to streamline operations and ensure consistency in its reporting.

## How It Works

The script performs the following steps:

1. **Login to Databases**: Automatically logs into the Zeniyacho and Fukune databases using provided credentials.
2. **Data Extraction**: Extracts sales data for the current period and the next period from both databases.
3. **Data Processing**: Processes the extracted data to calculate total earnings and perform various analyses.
4. **Report Generation**: Generates an Excel report that includes:
    - Total earnings for the current and next periods.
    - Comparison with the previous week’s earnings.
    - Highlighted changes in earnings.
5. **Report Storage**: Saves the generated report in the designated directory (`earnings_report_kyoto_machiyas`).

## Benefits to the Company

The automation of the earnings report generation brought several key benefits to the company:

1. **Time Savings**: Reduced the time required to compile weekly earnings reports from several hours to just a few minutes.
2. **Increased Accuracy**: Minimized human errors in data extraction and calculation, ensuring more reliable reports.
3. **Consistency**: Provided a standardized format for earnings reports, making it easier to track performance over time.
4. **Scalability**: Enabled the company to handle an increasing volume of data without a corresponding increase in workload.

## Future Enhancements

While the current version of the script has greatly improved the reporting process, there are several areas for future enhancements:

1. **Data Visualization**: Incorporate charts and graphs into the Excel report for better data visualization.
2. **Email Notifications**: Automatically send the generated reports via email to designated stakeholders.
3. **Real-time Reporting**: Develop real-time reporting capabilities to provide up-to-the-minute sales data.

## Acknowledgements

This project was made possible through the collaborative efforts of the data management and IT teams at Kyoto Machiyas. Special thanks to everyone who provided insights and feedback during the development process.
