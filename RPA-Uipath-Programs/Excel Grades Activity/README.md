# UiPath Excel Data Processing Activities

## Description
This project demonstrates Excel automation using UiPath Studio. It reads student marks from an Excel file, calculates the total, percentage and grade for every student, and saves the result to a new Excel file.

## Activities Used
- **Excel Application Scope:** Opens the Excel workbook.
- **Read Range:** Reads the student marks into a DataTable.
- **Add Data Column:** Adds the Total, Percentage and Grade columns.
- **For Each Row in Data Table:** Repeats the calculation for every student.
- **Assign:** Calculates the total and percentage.
- **If Activity:** Chooses the grade (A+, A, B, C, D or F) from the percentage.
- **Copy File:** Creates the result workbook from the original.
- **Write Range:** Writes the DataTable to a new sheet named Result.
- **Log Message and Message Box:** Report that the work is finished.

## Workflow
1. Read `students.xlsx` (sheet `Sheet1`) into the DataTable `dtStudents`.
2. Add the columns Total, Percentage and Grade.
3. For each student: Total = Maths + Science + English, Percentage = Total / 300 x 100.
4. Check the percentage with nested If activities: 90 and above is A+, 80 and above is A, 70 and above is B, 60 and above is C, 50 and above is D, otherwise F.
5. Copy `students.xlsx` to `students_result.xlsx`.
6. Write the DataTable to the sheet `Result` of `students_result.xlsx`.
7. Show a message that the result has been saved.

## Files
| File | Purpose |
|---|---|
| `Main.xaml` | The Sequence workflow. |
| `students.xlsx` | Input: Name, Maths, Science, English for 15 students. |
| `students_result.xlsx` | Output, created when the bot runs. |

## Expected Output
Sheet `Result` in `students_result.xlsx`:

| Name | Maths | Science | English | Total | Percentage | Grade |
|---|---|---|---|---|---|---|
| Aarav Sharma | 92 | 88 | 95 | 275 | 91.67 | A+ |
| Priya Nair | 85 | 90 | 78 | 253 | 84.33 | A |
| Rohan Gupta | 72 | 65 | 70 | 207 | 69.00 | C |
| Ananya Reddy | 95 | 97 | 92 | 284 | 94.67 | A+ |
| Vikram Singh | 55 | 48 | 60 | 163 | 54.33 | D |
| Sneha Patel | 78 | 82 | 74 | 234 | 78.00 | B |
| Karthik Iyer | 66 | 71 | 69 | 206 | 68.67 | C |
| Meera Joshi | 88 | 76 | 81 | 245 | 81.67 | A |
| Arjun Menon | 40 | 35 | 52 | 127 | 42.33 | F |
| Divya Krishnan | 91 | 84 | 89 | 264 | 88.00 | A |
| Rahul Verma | 62 | 58 | 67 | 187 | 62.33 | C |
| Isha Kapoor | 74 | 79 | 83 | 236 | 78.67 | B |
| Siddharth Rao | 50 | 45 | 58 | 153 | 51.00 | D |
| Neha Bhat | 83 | 92 | 87 | 262 | 87.33 | A |
| Manish Kumar | 35 | 42 | 47 | 124 | 41.33 | F |

## Tools Used
- UiPath Studio
- RPA
- Sequence Workflow
- Microsoft Excel and PowerPoint

## Objective
To automate reading, calculating and writing Excel data using UiPath, and to understand how a DataTable, loops and conditions work together in an RPA workflow.
