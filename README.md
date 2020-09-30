# STOCK ANALYSIS WITH VBA + EXCEL

## Overview of Project

### Purpose
A deep dive into Excel, as we know, Excel is a tool that can be used across all Organizations areas, including from household budgeting to complex financial analysis. 
Learning from the intricacies of Excel will draw on (and enhance) skills we may have already, like computer literacy, data literacy, and quantitative reasoning. 
Including advanced Excel features formulas, charts, and pivot tables. 

## Analysis and Challenges
Here's a quick look at the Kickstarting Analysis and Challenges of this Project, including the following tasks:

- Import data into a table for analysis.
- Apply filters, conditional formatting, and formulas.
- Generate and interpret pivot tables.
- Calculate summary statistics such as measures of central tendency, standard deviation, and variance.
- Characterize data to identify outliers in datasets.
- Perform an Excel analysis with visualizations.
- Interpret common Excel visualizations

#### Our Challenge Data Background
> Louise’s play Fever came close to its fundraising goal in a short amount of time. Now, she wants to know how different campaigns fared in relation to their launch dates and their funding goals. Using the Kickstarter dataset that you’ve already combed through, you’ll visualize campaign outcomes based on their launch dates and their funding goals. You’ll then submit a written report based on your analysis and the visualizations you create.

### Analysis of Outcomes Based on Launch Date
 
#### Deliverable Requirements with detail analysis:
**1. A Years column is created based on the Date Created Conversion column in the Kickstarter spreadsheet.**

> In the "Years" column, use the `YEAR()` function to extract the year from the “Date Created Conversion” column.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/A%20Years%20column%20is%20created.PNG?raw=true)


**2. A pivot table is created in a new worksheet labeled "Outcomes Based on Launch Date".**

> Created a pivot table from the KickStarter worksheet, and placed the pivot table in a new sheet.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Pivot%20Table%20for%20Outcomes%20Based%20on%20Launch%20Date.PNG?raw=true)


**3. The pivot table filters on "Parent Category" and "Years".**

> Placed the appropriate pivot table pivot table based on Parent Category and the Years data filtered.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Pivot%20table%20filters%20on%20Parent%20Category%20and%20Years.PNG?raw=true)


**4. The columns, rows, and values in the pivot table fields are correctly populated.**

> Placed the appropriate pivot table fields.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/A%20Years%20column%20is%20created.PNG?raw=true)


**5. The "Parent Category" is filtered on "theater".**

> Placed the appropriate filter on pivot table.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Filtered%20on%20Theater.PNG?raw=true)


**6. The row labels are changed to display the months of the year, and the campaign outcomes are sorted in descending order.**

> Grouping data in a PivotTable can help you show a subset of data to analyze. For example, you may want to group an unwieldy list of dates or times (date and time fields in the PivotTable) into quarters and months, etc.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Month%20on%20Row%20and%20Outcomes%20in%20descending.PNG?raw=true)


**7. A line chart is created showing the number of successful, failed, or canceled projects by month, it has a title, and it is saved as** **[Theater_Outcomes_vs_Launch.png]**

> Created a line chart from the pivot table to visualize the relationship between outcomes and launch month.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Theater_Outcomes_vs_Launch.PNG?raw=true)



### Analysis of Outcomes Based on Goals

#### Deliverable Requirements with detail analysis:
**1. A new sheet is created with eight columns and twelve rows, according to the instructions.**

> In the new sheet, create the following columns to hold the data:
> - Goal
> - Number Successful
> - Number Failed
> - Number Canceled
> - Total Projects
> - Percentage Successful
> - Percentage Failed
> - Percentage Canceled.

> In the “Goal” column, create the following dollar-amount ranges so projects can be grouped based on their goal amount.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/eight%20columns%20and%20twelve%20rows.PNG?raw=true)


**2. The `COUNTIFS()` function is used to populate the "Number Successful," "Number Failed," and "Number Canceled" columns, based on the project "outcome," the "goal" amount using the goal ranges in Step 3, and the Subcategory "plays".**

> Used `COUNTIFS()` functions to populate the "Number Successful," "Number Failed," and "Number Canceled" columns by filtering on the Kickstarter "outcome" column, on the "goal" amount column using the ranges created, and on the "Subcategory" column using "plays" as the criteria.Created a pivot table from the KickStarter worksheet, and placed the pivot table in a new sheet.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/The%20COUNTIFS()%20function.PNG?raw=true)


**3. The `SUM()` function is used on each row to add the "Number Successful," "Number Failed," and "Number Canceled" columns to populate the "Total Projects" column.**

> Use the `SUM()` function to populate the "Total Projects" column with the number of successful, failed, and canceled projects for each row.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/The%20SUM()%20function.PNG?raw=true)


**4. The percentages of successful, failed, and canceled projects are calculated based on the data from the "Total Projects," "Number Successful," "Number Failed," and "Number Canceled" columns.**

> Calculated the percentage of successful, failed, and canceled projects for each row.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/The%20percentages.PNG?raw=true)


**5. A line chart is created and saved as **[Outcomes_vs_Goals.png]** with the goal-amount ranges on the x-axis, the percentage of successful, failed, or canceled projects on the y-axis, and an appropriate title.**

> Created a line chart titled "Outcomes Based on Goal" to visualize the relationship between the goal-amount ranges on the x-axis and the percentage of successful, failed, or canceled projects on the y-axis.


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Outcomes_vs_Goals.PNG?raw=true)



### Challenges and Difficulties Encountered

Biggest challenge was filtering the pivot table to visualize the relationship between parent category and years, adding the correct dataset into the Columns, Rows and Values.


## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?

> As Conclusions, our Line charts we can see by looking at our data that the months of **May and June** both have a greater success rate.
> A bar chart **wouldn't** be able to convey this information in the same manner.



![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Theater_Outcomes_vs_Launch.PNG?raw=true)



- What can you conclude about the Outcomes based on Goals?

> As Conclusion, our Outcomes based on Goals measures using line chart of central tendency work in practice help us finding the mean and median for each dataset's (the failed and successful campaigns).


![name-of-you-image](https://github.com/emmanuelmartinezs/kickstarter-analysis/blob/master/artifacts_images/Outcomes_vs_Goals.PNG?raw=true)


- What are some limitations of this dataset?

> Some limitation can be that we'd like to know the deviations from the actual dataset, but because we don't know, these deviations have a subtle and slight bias to them. 


- What are some other possible tables and/or graphs that we could create?

> - Box Plots
> - Pie Graph
> - Column Graph
> - Line Graph
> - Area Graph
> - Scatter Graph




