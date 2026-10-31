# University Student Analytics Dashboard (Excel)

An interactive Excel dashboard I built to track student performance, attendance, fees and placements for a university with 1,000 students across three departments: Engineering, Science and Business Administration.

Everything is done inside Excel using Pivot Tables, Pivot Charts, Slicers and a few formulas. No Power BI and no VBA.

![Dashboard](images/dashboard.png)

## What the dashboard shows

The top row has four KPI cards:

- Total Students: 1,000
- Average Attendance: 89.35%
- Total Fees Collected: ₹4.38 Cr
- Students Placed: 489

Below that there are five charts:

- Top 10 Students by CGPA
- Students by Department
- Average Percentage by Program
- Pass vs Fail split
- Placement Status

The slicers on the left (Gender, Department, Program and Year) are connected to all the pivot tables, so every chart and KPI card updates together. For example, you can filter down to only female B.Tech students in 3rd year and the whole dashboard changes.

## Workbook structure

| Sheet | What's in it |
|---|---|
| Dashboard | The final dashboard with KPI cards, charts and slicers |
| Dataset | Main table (Table1), 1,000 students and 35 columns: personal details, department, program, year, semester, attendance, CGPA, fees, placement, result and subject marks |
| Attendance | Month-wise attendance from January to December, plus the overall % |
| Marks | Subject-wise marks for each student (5 subjects, each out of 100) |
| KPIs and Charts | All the pivot tables the dashboard is built on |

Each program has its own 5 subjects:

- B.Sc: Biology, Chemistry, Mathematics, Environmental Science, Physics
- B.Tech: Physics, Electronics, Engineering Drawing, Maths, Programming
- MBA: Business Analytics, Operations, Finance, HRM, Marketing

## How I built it

1. Converted the Dataset sheet into an Excel Table so the pivot tables pick up new rows automatically.
2. Filled the Program column from Department using an `IFS` formula (Engineering = B.Tech, Science = B.Sc, Business Administration = MBA).
3. Calculated `Marks Obtained` as the sum of the subject columns and `Percentage` as Marks Obtained / 500 * 100.
4. Pulled `Attendance %` into the Dataset from the Attendance sheet with `INDEX` / `MATCH`, so attendance only comes from one place.
5. Made the pivot tables in the KPIs and Charts sheet. The KPI cards on the dashboard are text boxes linked to `GETPIVOTDATA` cells, so they change with the slicers too.
6. Created pivot charts from those pivots and connected all the slicers through Report Connections.
7. Put it all together on the Dashboard sheet with the university logo and a common colour theme.

## Fixes I made after checking the numbers

When I cross-checked the dashboard against the raw sheets, a few things were off. This is what I changed:

- **Subject marks were matched to the wrong students.** The Marks sheet is sorted by program, and the marks had been pasted into the Dataset in row order instead of by Student ID, so for example a B.Tech student had B.Sc subjects. I re-matched every row on Student ID, which also corrected Marks Obtained and Percentage for each student.
- **Two different attendance numbers.** The Dataset said 75% average and the Attendance sheet said 89%. The month-wise sheet is the detailed source, so the Dataset now looks it up from there.
- **Top 10 chart was adding CGPAs together.** Two names appear twice in the data (Manikya Malhotra and Bhamini Sodhi), and because the pivot used Sum, they showed CGPAs of 13.93 and 12.48. Changed it to Max, so it now shows the real top 10.
- **"No. of Admissions" chart was actually students per department.** Renamed it to Students by Department and changed it from a line chart to a column chart, since departments are categories and not a timeline.
- **Average Marks chart had two scales on one axis.** Marks out of 500 and percentage were plotted together, so the percentage bars were tiny. It now shows only the average percentage.
- **Fees card was showing "₹ 44.M".** That was a number format problem. It now shows the amount in crores.
- **Year and semester didn't match**, e.g. Year 1 with Semester 7. Semester now follows the year (Year 1 = Sem 1 or 2, Year 2 = Sem 3 or 4, and so on). MBA is a 2-year course, so MBA students marked as Year 3 or 4 were moved to Year 1 or 2.
- Fixed the `Studen _ID` header typo to `Student_ID` and removed some leftover test text from the dashboard sheet.

## Some things I noticed in the data

- Engineering is about two-thirds of the university (664 out of 1,000 students). Business Administration has 182 and Science has 154.
- Attendance is almost the same everywhere, between 89.2% and 89.9% for all three departments.
- Only about half the fees have come in so far: ₹4.38 Cr collected out of ₹8.71 Cr billed (50.3%). Science is the lowest at 47.8%.
- MBA has the best placement rate (54.9%) but the lowest pass rate (44.5%). B.Sc is the other way round, with 42.9% placed and 51.9% passing.
- Average percentage is close to 70% for all three programs, so no program is clearly ahead on marks.
- Placement doesn't seem to depend on CGPA here. Roughly half the students are placed in every CGPA range, which is most likely because this is practice data.

## Limitations

- This is a practice dataset. The names and numbers are made up and are not real students.
- Placement status isn't linked to year in the data, so some 1st year students show up as placed. I left this as it is.
- Slicers need desktop Excel (2010 or later). They won't work properly in Google Sheets or LibreOffice.

## How to use it

1. Download `Student_Analytics_Dashboard.xlsx` or clone the repo.
2. Open it in desktop Excel. If it opens in Protected View, click **Enable Editing** so the pivot tables refresh.
3. Click on the slicers to filter. Use the clear filter button on the top right of each slicer to reset it.

If you just want to look at the data without downloading anything, there are CSV copies of the three data sheets in the `data` folder, which GitHub shows as tables.

## Repo contents

```
├── Student_Analytics_Dashboard.xlsx
├── README.md
├── images/
│   └── dashboard.png
└── data/
    ├── students.csv
    ├── attendance.csv
    └── marks.csv
```

## Tools used

Microsoft Excel: Pivot Tables, Pivot Charts, Slicers, `GETPIVOTDATA`, `IFS`, `INDEX` / `MATCH`, Excel Tables and conditional formatting.

## Author

Harsh
