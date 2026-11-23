# UiPath Excel to PowerPoint Activities

## Description
This project demonstrates Excel and PowerPoint automation using UiPath Studio. It reads sales data from an Excel file and fills a PowerPoint template to create a sales report automatically.

## Activities Used
- **Excel Application Scope:** Opens the Excel workbook.
- **Read Range:** Reads the sales data into a DataTable.
- **For Each Row in Data Table:** Goes through every product.
- **Assign:** Builds the product list and adds up the total units and total revenue.
- **Use PowerPoint Presentation:** Opens the PowerPoint template.
- **Add Text to Slide:** Writes text into named shapes on the slides.
- **Save PowerPoint File As:** Saves the finished report as `Sales_Report.pptx`.
- **Log Message and Message Box:** Report that the work is finished.

## Workflow
1. Read `sales.xlsx` (sheet `Sheet1`) into the DataTable `dtSales`.
2. For each product, add a line (product, units, revenue) to the variable `productList` and add to the totals.
3. Open `sales_template.pptx`.
4. Add the date to slide 1 (shape `DateText`), the product list to slide 2 (shape `ProductList`), and the totals to slide 3 (shapes `TotalUnitsText` and `TotalRevenueText`).
5. Save the presentation as `Sales_Report.pptx` (the template is not changed).
6. Show a message that the report has been created.

## Files
| File | Purpose |
|---|---|
| `Main.xaml` | The Sequence workflow. |
| `sales.xlsx` | Input: Product, Units Sold, Revenue for 8 products (revenue in INR). |
| `sales_template.pptx` | Template with 3 slides and named shapes. |
| `Sales_Report.pptx` | Output, created when the bot runs. |

## Expected Output
`Sales_Report.pptx` with 3 slides:
1. **Sales Report** with today's date.
2. **Product-wise Sales**:
```
Laptop   |   45 units   |   Rs. 2,925,000
Smartphone   |   150 units   |   Rs. 2,700,000
Tablet   |   80 units   |   Rs. 1,760,000
Smartwatch   |   95 units   |   Rs. 807,500
Headphones   |   220 units   |   Rs. 550,000
Keyboard   |   310 units   |   Rs. 372,000
Monitor   |   60 units   |   Rs. 840,000
Webcam   |   140 units   |   Rs. 448,000
```
3. **Summary**: total units sold **1,100** and total revenue **Rs. 10,402,500**. (Number grouping follows the PC's regional settings.)

## Tools Used
- UiPath Studio
- RPA
- Sequence Workflow
- Microsoft Excel and PowerPoint

## Objective
To automate the creation of a PowerPoint report from Excel data using UiPath, removing manual copying and typing.
