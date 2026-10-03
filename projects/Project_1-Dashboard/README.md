# Excel Salary Dashboard

![1_Salary_Dashboard.png](/0_Resources/Images/1_Salary_Dashboard_Final_Dashboard.gif)

## Introduction

This project is an interactive Excel dashboard designed to analyze salary trends across data-related job roles. It allows users to explore median salaries based on job title, location, and job schedule type.

The dashboard uses job-market data containing information about job titles, salaries, locations, and relevant skills to provide a clear view of salary patterns across different roles and regions.

### Dashboard File
My final dashboard is in [1_Salary_Dashboard.xlsx](1_Salary_Dashboard.xlsx).

### Excel Skills Used

The following Excel features were used to build and analyze the dashboard:
- **📉 Charts
- **🧮 Formulas and Functions
- **❎ Data Validation

### Data Jobs Dataset

The dataset used in this project contains data science and data-related job information from 2023. It includes detailed information on:

- **👨‍💼 Job titles**
- **💰 Salaries**
- **📍 Locations**
- **🛠️ Skills**

## Dashboard Build

### 📉 Charts

#### 📊 Data Science Job Salaries - Bar Chart

<img src="/0_Resources/Images/1_Salary_Dashboard_Chart1.png" width="850" height="550" alt="Salary Dashboard Chart1">

- 🛠️ **Excel Features:** Used a horizontal bar chart with formatted salary values to present the data clearly.
- 🎨 **Design Choice:** A horizontal bar chart was selected to make salary comparisons between job titles easier to read.
- 📉 **Data Organization:** Job titles were sorted by median salary in descending order.
- 💡 **Insights Gained:** The visualization makes it easy to compare salary levels across different data-related roles and identify higher- and lower-paying positions.

#### 🗺️ Country Median Salaries - Map Chart

![1_Salary_Dashboard_Chart2.png](/0_Resources/Images/1_Salary_Dashboard_Country_Map.gif)  

The dashboard also includes a map chart showing median salaries by country.

- 🛠️ **Excel Features:** Used Excel's Map Chart feature to visualize median salaries geographically.
- 🎨 **Design Choice:** Color-based geographic visualization makes differences in salary levels easier to identify.
- 📊 **Data Representation:** Median salary was plotted for each country with available data.
- 👁️ **Visual Enhancement:** The map provides a quick overview of geographic salary trends.
- 💡 **Insights Gained:** The visualization highlights differences in median salaries across countries and regions.

### 🧮 Formulas and Functions

#### 💰 Median Salary by Job Titles

```
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
```

- 🔍 **Multi-Criteria Filtering:** Checks job title, country, and schedule type while excluding records without a valid salary.
- 📊 **Array Formula:** Uses MEDIAN() together with IF() to calculate the median from the filtered data.
- 🎯 **Dynamic Analysis:** Allows salary information to be calculated according to the selected job title, country, and schedule type.
- **🔢 Formula Purpose:** The result is used to populate the salary table that supports the dashboard.

🍽️ Background Table

![1_Salary_Dashboard_Screenshot1.png](/0_Resources/Images/1_Salary_Dashboard_Screenshot1.png)

📉 Dashboard Implementation

<img src="/0_Resources/Images/1_Salary_Dashboard_Job_Title.png" width="400" height="500" alt="Salary Dashboard Title">

#### ⏰ Count of Job Schedule Type

```
=FILTER(J2#,(NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*(J2#<>0))
```

- 🔍 **Unique List Generation:** Uses FILTER() to remove entries containing "and" or commas and exclude zero values.
- **🔢 Formula Purpose:** Creates a cleaned list of job schedule types that can be used throughout the dashboard.

🍽️ Background Table

![1_Salary_Dashboard_Type.png](/0_Resources/Images/1_Salary_Dashboard_Screenshot2.png)

📉 Dashboard Implementation:

<img src="/0_Resources/Images/1_Salary_Dashboard_Type.png" width="350" height="500" alt="Salary Dashboard Type">

### ❎ Data Validation

#### 🔍 Filtered List

- 🔒 **Enhanced Data Validation:** Data validation was implemented for the Job Title, Country, and Type selections in the dashboard. This helps to:
    - 🎯 Restrict user input to predefined values
    - 🚫 Prevent incorrect or inconsistent entries
    - 👥 Improve the overall usability of the dashboard
    - 🔄 Make the dashboard easier to interact with and explore

<img src="/0_Resources/Images/1_Salary_Dashboard_Data_Validation.gif" width="425" height="400" alt="Salary Dashboard Data Validation">

## Conclusion

This project demonstrates how Excel can be used to transform job-market data into an interactive salary analysis dashboard.

By combining Excel formulas, charts, map visualizations, and data validation, I created a dashboard that allows users to explore salary trends across different data-related job titles, countries, and job schedule types.

The project also demonstrates practical skills in data analysis, visualization, and dashboard development using Excel.
