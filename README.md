# STOCK ANALYSIS WITH VBA + EXCEL

## Overview of Project

### Purpose
In this project and analyisis, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.. 

## Analysis and Challenges
Here's a quick look at the Kickstarting Analysis and Challenges of this Project, including the following tasks:

- Prepare our dataser `VBA_Challenge.vbs` file for the project.
- Create our resources folder in **GitHub** to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.
- Create and convert our `XLSM` file from `*.vbs` dataset that you used in this module as `VBA_Challenge.xlsm`.
- Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
- Use the steps **Refactor VBA code and measure performance** to add code where indicated by the numbered comments in the starter code file.

#### Our Challenge Data Background
> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

> In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

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




