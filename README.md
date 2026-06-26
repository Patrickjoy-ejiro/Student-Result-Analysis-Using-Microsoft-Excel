# Student Academic Performance Analysis Using Microsoft Excel

<img width="980" height="980" alt="17825073329598657699310071366495" src="https://github.com/user-attachments/assets/fa53afe0-8c25-4dcb-8dd3-aa14e4a6078f" />


## Project Overview
This project focuses on analyzing student academic performance using Microsoft Excel. The analysis transforms raw examination data into meaningful insights by evaluating overall performance, grade distribution, identifying students requiring academic support, and providing recommendations for improving future academic outcomes.

## Questions
1. Calculate the overall student performance.
2. Show how scores are distributed across grades.
3. Show at least top students that did excellently and those that needs help (using chart) 
4. Show the ranking/remark of each grade/students I'd. 
5. Provide data-driven recommendations to improve student performance.
   
## Tools Used
- Microsoft Excel
- Excel Formulas
- Pivot Tables
- Charts
- Slicers
- Dashboard

---

# Analysis and Test Result Calculations

## 1. Overall Student Performance

### Calculation Table

| Performance Indicator | Excel Formula | Result |
|---|---|---|
| Total Students | =COUNTA(Student_ID) | 56 |
| Class Average Score | =AVERAGE(Average) | 66.61% |
| Highest Score | =MAX(Average) | 97.17% |
| Lowest Score | =MIN(Average) | 30.25% |

### Interpretation
The overall class performance shows an average score of 66.61%. While some students achieved excellent results, some students recorded low scores and require improvement support.

---

## 2. Grade Distribution Analysis

### Grade Calculation

| Grade | Score Range | Number of Students |
|---|---|---|
| A | 80 - 100 | 19 |
| B | 70 - 79 | 6 |
| C | 60 - 69 | 13 |
| D | 50 - 59 | 6 |
| F | Below 50 | 12 |

### Excel Formula

```excel
=COUNTIF(Grade,"A")
### 4. Student Ranking Calculation

| Student | Average Score | Rank |
|---|---|---|
| Student 1 | 97.17% | 1 |
| Student 2 | 92.50% | 2 |
| Student 3 | 90.00% | 3 |

**Formula Used:**

`=RANK(Average_Score,Average_Range,0)`


---

### 5. Recommendations

- Provide extra lessons for students below 60%.
- Monitor student progress through continuous assessment.
- Encourage peer learning and revision practices.
- Focus teaching efforts on weak subject areas.
## Key Findings
- Evaluated the overall class performance.
- Identified students performing at different grade levels.
- Highlighted students who need additional academic support.
- Provided strategies for improving future examination results.

## Conclusion
This project demonstrates how Excel can be used to transform student records into meaningful insights that support educational decision-making and improve academic performance.
