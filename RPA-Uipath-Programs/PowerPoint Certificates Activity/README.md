# UiPath PowerPoint Certificate Generator

## Description
This project demonstrates PowerPoint automation using UiPath Studio. It creates one certificate for every participant listed in an Excel file by filling a PowerPoint template (a mail merge done with RPA).

## Activities Used
- **Excel Application Scope:** Opens the Excel workbook.
- **Read Range:** Reads the participants into a DataTable.
- **Assign:** Sets the output folder path.
- **Create Folder:** Creates the `Certificates` folder.
- **For Each Row in Data Table:** Goes through every participant.
- **Use PowerPoint Presentation:** Opens the certificate template.
- **Replace Text in Presentation:** Replaces `<<NAME>>`, `<<COURSE>>` and `<<DATE>>` with the participant's details.
- **Save PowerPoint File As:** Saves each certificate under the participant's name.
- **Log Message and Message Box:** Report progress and completion.

## Workflow
1. Read `participants.xlsx` (sheet `Sheet1`) into the DataTable `dtParticipants`.
2. Create the `Certificates` folder.
3. For each participant, open `certificate_template.pptx`.
4. Replace the three placeholders with the participant's name, course and date.
5. Save the file as `Certificates\Certificate_<Name>.pptx` (the template is not changed).
6. Show a message when all certificates have been created.

## Files
| File | Purpose |
|---|---|
| `Main.xaml` | The Sequence workflow. |
| `participants.xlsx` | Input: Name, Course, Date for 10 participants. |
| `certificate_template.pptx` | Template with the placeholders `<<NAME>>`, `<<COURSE>>` and `<<DATE>>`. |
| `Certificates/` | Output folder, created when the bot runs. |

## Expected Output
10 certificates in the `Certificates` folder:

- Certificate_Aarav_Sharma.pptx
- Certificate_Priya_Nair.pptx
- Certificate_Rohan_Gupta.pptx
- Certificate_Ananya_Reddy.pptx
- Certificate_Vikram_Singh.pptx
- Certificate_Sneha_Patel.pptx
- Certificate_Karthik_Iyer.pptx
- Certificate_Meera_Joshi.pptx
- Certificate_Arjun_Menon.pptx
- Certificate_Divya_Krishnan.pptx

## Tools Used
- UiPath Studio
- RPA
- Sequence Workflow
- Microsoft Excel and PowerPoint

## Objective
To automate repetitive document creation using UiPath by combining Excel data with a PowerPoint template.
