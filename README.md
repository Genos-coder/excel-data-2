# 📊 Excel for Data Analytics

## Table of Contents

### [Chapter 1: Spreadsheets_Intro](/1_Spreadsheets_Intro/)

- Worksheets
- Workbooks
- Ribbon
- Menu
- Keyboard Shortcuts

### [Chapter 2: Formulas & Functions](/2_Formulas_Functions/)

- Formulas
- Functions
- Logical Functions
- Text Functions
- Date Functions
- Lookup Functions
- Math Functions
- Statistical Functions

### [Chapter 3: Charts](/3_Charts_Graphs/)

- Chart Types
- Chart Elements
- Chart Styles
- Chart Axes
- Chart Layouts
- Chart Templates

### [Chapter 4: Spreadsheets Advanced](/4_Spreadsheets_Advanced/)

- Tables
- Conditional Formatting
- Formatting
- Collaboration

### [Chapter 5: Pivot Tables](/5_Pivot_Tables/)

- Pivot Tables
- Pivot Charts
- Pivot Tables Advanced

### [Chapter 6: Advanced Data Analysis](/6_Advanced_Data_Analysis/)

- Analysis Add-ins
- Solver
- Scenario Manager
- Goal Seek
- Data Tables

### [Chapter 7: Power Query](/7_Power_Query/)

- Power Query
- Power Query Editor
- Power Query M Language
- Power Query Advanced

### [Chapter 8: Power Pivot](/8_Power_Pivot/)

- Power Pivot
- Power Pivot DAX
- Power Pivot Data Models

## [Chart Statistics]()

#### Bar / Whiskers Chart

- Create the box plot using data from workbook 3_chart_statistics , worksheet -> Box plot_2

- Edit this graph make it's appearance visually appealing

- First change max value to `300000` by clicking on x-axis label then go to -> Axis options

- Change number format to example `200K`

#### Spark-lines

- It is nothing but way to insert mini charts in the row

- Select the Data from C4 to N10 and create the spark-line

- Add high and low point to it

- Change the marker color for high and low point

# Advance Spreadsheet

## Tables

- Go to insert and click on table to create the table or use short ctrl + T

- The table design tab get's enable click on it and change the table name to jobs

- Go to table style options

- If you add any column adjacent to that table it get's added to that table

- Table has their own formulas

- Create column salary_year_avg_copy and copy the content inside the salary_year_avg into it

- Create the column named excel and check that the excel skill is present in the job_skills column

- Add the total row using table design tab

- SUBTOTAL function vs AGGREGATE function

- Using SUBTOTAL function you can add formula number to it

- Using AGGREGATE function you can have one more argument which is options

```
=SUBTOTAL(function_num, row_ref)

=AGGREGATE(function_num, options, row_ref)
```

## Table Limitation

- Go to table_limits_original worksheet

- Convert the data into table using ctrl + T

- Filter the jobs based on data analyst

## Table slicers and & Combo

- Go to histogram_original worksheet and then click ctrl + T to convert the data into table

- Go to table design -> click on insert slicers -> Tick top three option and create the slicers

- use slicer to filter the data and observe the histogram

- Enable multiple selection using top right icon in the slicer

## Formatting

#### Normal Formatting

- Go to the 2_formatting excel workbook

- Clear the formatting using command tools in the home tab

- Format the table like headers, border, spacing

#### Conditional formatting

- Use conditional formatting and format the numeric values in the table

- clear the formatting using same tab

- Use format painter to apply same formatting

- Go to manage rules by clicking on the conditional formatting command

- Format the job count using data bar in conditional formatting

- Make it visually more appealing using manage rules

- Use star icon to show the job rank and remove the numbers using manage rules

# Getting Started With Our Project

- Go to the collaboration and create the new worksheet calculator

- Create validation worksheet and get the unique job_titles and sort them based on their count in the data worksheet

- For excel 2019 we cannot use unique function to get the unique job titles

#### Algorithm to find unique values from one column

- step1: Create the helper column where you get the row number of the job title which occur first time
- step2: Use if condition and pass `COUNTIF` and provide range of first cell to that same cell and fix the cell reference of first cell in range example `$A$2:A2` so that when we autofill it will check all the cells from start to check the A2 appear how many times
- step3: As we know we want unique values the above method should return only 1 count and based on that we will return the row index using `ROW()` in if condition else we will return "" empty cell
- step4: Now go to the data validation table and create the column with header job_title_short use `MATCH()` to find the index of the small row number from that helper function and use `SMALL()` which takes range and which small number you want from that range starting from first to number you provide then match that lookup number using match function to get cell reference of that number from the range
- step5: Once you got the cell reference we can use `INDEX()` method to get the value from job_title_short column which are unique

#### Algorithm to sort the job_titles

- we have to sort the values based on the count of job_titles
- step1: create the helper column which rank the all count starting from 1 to n
- step2: Use that helper column and find the cell reference using match function and the lookup value in the match function should start from 1
- step3: after getting the cell reference find the value by passing it inside the index function

#### Task: Sort count using above algorithm

## Creating data validation

- Go to calculator worksheet and create data validation in front of the job title cell
- Go to data tab -> in allow: click on dropdown and select list because we are adding the job_title there

- As we done this the values other than the list we provide will not get accepted

- We are doing this to calculate the median salary based on the values in that data validated cell

## Finding the median salary based on the value in data validation dropdown we created

- Create new worksheet named salary
- Copy the sorted job title from it
- Find the median salary using custom median formula

- now sort it based on the median salary in descending order

- Create the cell where you can get the median salary based on the value you from the data validation dropdown you have created

## Protecting data or your work from co-workers

- select all cells

- Unselect the data validated cell which has the dropdown

- Got to review tab -> select protect the sheet -> and unclick the select the locked cells

- Click on the worksheet and hide it and if you want to un-hide the worksheet you can right click on it and select the worksheet name to un-hide it

# Creating the dashboard of our project and formatting the all stuff

- Place the content as per our dashboard

## Creating country type

- Create worksheet name country

- Create the column country in the data validation worksheet and get the unique country names and sort them alphabetically

- Create the copy of country and paste it as special -> values

- select all country copy -> go to data tab -> sort it A->Z

- Now go to calculator worksheet and create country header and under that create data validation and add those sorted countries

## Job Schedule type

- create helper column and find the unique job schedule type in validation worksheet

- Filter down them to single job type

- Take the filtered value into another column

- create the header type and add dropdown using data validation and filtered job schedule type

## Median salary according to country

- Create the country worksheet copy paste the unique job countries

- get the median salary according to them

- also filter the median salary based on the job title we select

## NOTE: You can check the names you provided to the cell by going to the formulas tab -> name manager or (ctrl+F3)

- Add one more filter while calculating the median salary according to the title we select

- Also also filter the median salary according to type

- Make more precision

```
=MEDIAN(
     IF(
          (jobs[job_country]=A2)*
          (jobs[salary_year_avg]<>0)*
          (jobs[job_title_short]=title)*
          (ISNUMBER(SEARCH(type,jobs[job_schedule_type]))),jobs[salary_year_avg]))
```

- Now make the column of those country who has median salary not a num error and sort them from highest to lowest

- Create the map using country and median salary column

## Manipulating the graphs based on job title
