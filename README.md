# Student Academic Performance Analysis Using Microsoft Excel

<img width="1280" height="720" alt="image" src="https://github.com/user-attachments/assets/59d9741d-6a93-4730-9d5e-5a08c21615be" />


## Introduction

This project analyzes student examination results using Microsoft Excel data analysis tools. The purpose is to evaluate overall student performance, understand grade distribution, identify students who require academic support, and provide recommendations for improving future academic outcomes.

---

## Questions

1. Calculate the overall student   performance.
2. Show how scores are distributed across grades.
3. Identify top-performing students and students needing support.
4. Show ranking and remarks of each student.
5. Provide data-driven recommendations to improve student performance.

---

## Tools Used

- Microsoft Excel
- Excel Formulas
- Pivot Tables
- Charts
- Slicers
- Dashboard
- Power Bi

---

# 1. Overall Student Performance

## Performance Summary

| Performance Indicator | Formula | Result |
|---|---|---|
| Total Students | =COUNTA(Student_ID) | 56 |
| Class Average Score | =AVERAGE(Average) | 66.61% |
| Highest Score | =MAX(Average) | 97.17% |
| Lowest Score | =MIN(Average) | 30.25% |

### Interpretation

The overall class performance shows an average score of 66.61%. The result indicates a fair academic performance, although some students require additional academic support to improve their results.

---

# 2. Grade Distribution Analysis

## Grade Summary Table

| Grade | Score Range | Number of Students |
|---|---|---|
| A | 80 - 100 | 19 |
| B | 70 - 79 | 6 |
| C | 60 - 69 | 13 |
| D | 50 - 59 | 6 |
| F | Below 50 | 12 |

## Excel Formula

```excel
=COUNTIF(Grade,"A")
```

### Chart Used

**Pie Chart**

Purpose:
- Shows the number of students in each grade category.
- Helps compare performance levels.

---

# 3. Top Students and Students Needing Support

## Top Performing Students

| Student ID | Average Score | Grade | Remark |
|---|---|---|---|
| Student 1 | 97.17% | A | Excellent |
| Student 2 | 92% | A | Excellent |
| Student 3 | 90% | A | Excellent |

## Students Needing Support

| Student ID | Average Score | Grade | Remark |
|---|---|---|---|
| Student 10 | 45% | F | Needs Support |
| Student 15 | 50% | D | Needs Support |
| Student 20 | 55% | D | Needs Support |

### Chart Used

**Bar Chart**

Purpose:
- Compares excellent students and students requiring intervention.

---

# 4. Student Ranking and Remarks

## Ranking Table

| Student ID | Average Score | Rank | Remark |
|---|---|---|---|
| Student 1 | 97.17% | 1 | Excellent |
| Student 2 | 92% | 2 | Excellent |
| Student 3 | 90% | 3 | Excellent |

## Rank Formula

```excel
=RANK(Average,Average_Range,0)
```

## Remark Formula

```excel
=IF(Average>=80,"Excellent",IF(Average>=60,"Good",IF(Average>=50,"Average","Needs Support")))
```

---

# 5. Data-Driven Recommendations

Based on the analysis:

- Provide extra lessons for students scoring below 60%.
- Organize revision classes and academic support programs.
- Encourage peer learning and group study.
- Monitor student progress through continuous assessment.
- Identify weak subjects and improve teaching strategies.
- Provide advanced learning activities for high-performing students.

---

# Conclusion

This project demonstrates how Microsoft Excel can be used to analyze student academic results, visualize performance trends, identify learning gaps, and support educational decisions through data-driven insights. Also how Power Bi is used for a perfect Data Visualization. 

I'm interested in collaborating with projects that involves analysis.

You can reach me: (ejirolink@gmail.com)
