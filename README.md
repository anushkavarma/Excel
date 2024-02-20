# Project 1 Elements
I polished my Microsoft Excel skills for topics including Pivot Tables, a variety of formulas, XLOOKUP, conditional formatting, charts, and cleaning data. All of these can be found as files in this repository and I have shown the formulas used and the explanations for what my goal for each of these were, such as finding the corresponding email for an employee.

I also completed a project in Excel for bike sales, including cleaning data, clarifying and simplifying the data, and creating pivot tables, charts, and a dashboard to visualize the data. I have also logged the steps taken, formulas and functions used, and the purpose and result of each of my steps.

### Pivot Tables
1. Taking bike sales data, I developed a pivot table of revenue per country broken down by state that are in collapsible fields. This pivot table showed that the United States brought in the highest revenue from bike sales.

2. Next, I developed a pivot table with values for revenue, costs, and profits from bike sales. I created a Calculated Field to confirm the calculations for the sum of profit using the Pivot Table Analyze function.

3. The third pivot table shows the revenue with a filter for gender. From this, the data can be used to determine whether male or female customers should be a focus for future sales.

4. The Revenue Per Year pivot table is divided by country and shows the countries and years that performed best in sales -- on a large scale, each country generally had an increase in year-to-year sales. This pivot table helps to make the many rows of data understandable at a large scale.

### Formulas
For Employee data, I used **Max and Min** formulas to find the latest and earliest start dates and highest and lowest salaries among the employees.

I used an **IF statement** to calculate whether an employee was older or younger than 30, and name them "Old" or "Young" accordingly. I also used **IFS statements** to find each employee's job title and assign them to a department such as Sales.

I used the **LEN** formula to determine whether employee phone numbers were valid. If the length/LEN formula presented a character length that was not equal to 10, then my following IF statement would present a result of "invalid."

The **TEXT** formula helped me to convert date to text and then, I could use the **Left and Right** formulas on the dates to find the start year for each employee.

To generate emails for each employee, I used the **Concatenate** formula to create an email of FirstName.LastName@dm.com.

The **Trim** formula removed any unnecessary spaces before and after employee names to clen the data.

The **Substitute** formula helped to standardize the start dates in different formats using 1, 2, and no instances.

I used **SUM and SUM IFS statements** to calculate the sum of employee salaries if the employee was 30+ and female, and the **COUNT and COUNT IFS statements** to calculate the sum of salaries if the employee had a certain employee number and was male.

Finally, I used **DAYS and NETWORKDAYS** to calculate the number of days and working days worked by each employee based on their start and end date.

### XLOOKUP
I first used **XLOOKUP** to find the employee email for the corresponding name. I also used this formula for multiple rows at once, to find the end date and email for each employee.

Since some of the data was incomplete for employee names, I used **wildcard character match** to find employee emails.

In order to find the person who had the closest start date to a certain date, I used **search order** and **match order**.

For sales of the company's products, I used **XLOOKUP horizontal search** to find the corresponding paper sales value, and to find the sales for multiple months, I used **XLOOKUP with sum** to add the February and March sales values.

Finally, I used **VLOOKUP** to search the table array and find the email addresses for corresponding employee.

### Conditional Formatting
Conditional formatting can be useful in showing patterns and trends.

To show how the products were selling from month to month, I used **icon sets** to show directional changes, to organize salaries from highest to lowest, I used a **color scale**, and to highlight employee salaries that are above average, I used **top/bottom rules**.

To find any duplicates and incorrect formatting in data, highlighting this data can be useful to detect it, so I used **highlight cells rules** to find any duplicated start dates and incorrect date formats, and sorted the data to show the duplicates at the top.

I also **created a new rule** to highlight employees who are older than 30, and those with a salary higher than $50,000.

### Charts
I created a bar chart of products sold per month, a line graph with sales per month for certain products, and I also added a data table with a legend to show the sales for each product each month.

### Cleaning Data
For a dataset about presidents, I cleaned the data to **remove duplicates**, using **Upper** and **Proper** to fix capitalized names, **Find and Replace** to correct issues in party names, **Trim** to remove unnecessary spaces, and changed **date formats** to standardize all date formats.

# Full Project 1
The goal of this project was to take data, clean it, and create an interactive dashboard using that data. The dataset contained bike sales customer data. For this project, I included a log of the steps I took in the Excel file.

### Cleaning the data
I **removed duplicates**, used **find and replace** to clarify data such as marital status and gender to change letters including M, S, and F into Married, Single, Male, and Female.

I them used an **IFS statement** to create a new column for Age Brackets in order to condense the information and make it easier to understand. Based on whether a customer was <31, 31-54, or 54+, they were categorized as young adult, middle age, and senior, respectively.

### Pivot Tables
These pivot tables were used to make the dashboard.

1. The purpose of the first pivot table was to look at the average income of people who did or did not buy a bike, to see whether income plays a role in whether someone makes a purchase. This can determine who should be targeted as potential customers. The pivot table contained information of income, gender, and bike purchase (yes/no). I created a clustered column chart to visualize the data, which showed that males had a higher income and purchased bikes more, which suggests a correlation between the two.

2. The second pivot table looked at commuting distance of customers and whether they purchased a bike, to see if customers live relatively close or far to their workplace. The pivot table contained the fields of commute distance, purchased bike (yes/no), and purchased bike (count). I then created a line graph to visualize the customer commute data, which illustrated that the largest number of customers lived 0-1 miles from work.

3. The third pivot table focused on customer age brackets and whether they purchased a bike or not. The fields included age brackets, purchased bike (count), and purchased bike (yes/no). I developed a line graph from this pivot table that shows that middle-aged people are most likely to purchase a bike. Using age brackets here shows a clean result rather than individual ages.

### Dashboard
I designed a dashboard and then included **slicers** for marital status, region, and education applied to all charts in order to see how the results change based on certain demographic information, with the ability to see this data visualized in charts on the dashboard.
