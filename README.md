# Working with Excel Using PowerShell

**Author: Longtao Xu**

## Introduction

This lab helped me practise using PowerShell to read and create Excel files. I used the ImportExcel module on a Windows Server virtual machine and opened the finished workbook in WPS Office on my host computer.

I also learned how *Markdown* can organize technical notes in a Git repository. Clear headings, lists, tables, and inline code make the steps easier to follow.

## Preparing the Environment

Installing ImportExcel directly on the server failed. The connection tests showed that the server could not resolve the website names needed for the download.

I downloaded the module on the host computer, copied it to the virtual machine through PowerShell Direct, and loaded it with `Import-Module ImportExcel`. The module version shown in my session was **7.8.10**.

### Tools and Files

- PowerShell
    - Run commands and inspect objects.
    - Transfer files between the host and virtual machine.
- ImportExcel
    - Import spreadsheet data.
    - Export objects and create charts.
- WPS Office
    - View the completed workbook.
- GitHub
    - Store this Markdown article in a Git repository.

## Reading and Exporting Data

I imported the supplied `Fruits.xlsx` file using `Import-Excel` with `-OutVariable fruit`. This displayed the data and stored the imported results in `$fruit`.

Running `$fruit | Get-Member` showed that the imported records were `PSCustomObject` objects. The two data properties were `Fruits` and `Quantity`.

### Fruit Data

| Fruit | Checked | Quantity |
| :--- | :---: | ---: |
| Bananas | Yes | 12 |
| Blueberries | Yes | 18 |
| Apples | Yes | 2 |
| Oranges | Yes | 11 |
| Pears | Yes | 15 |
| Cherries | Yes | 9 |

Blueberries had the highest quantity, while apples had the lowest. The column chart displayed the same values as the original data.

### Process Information

I exported the output of `Get-Process` to `ProcessInfo.xlsx` and named the worksheet `Processes`.

The column numbers in the lab example selected different properties in my file. I inspected the headers and changed the selection to `@(1,48,10,6)`. This returned `Name`, `ProcessName`, `CPU`, and `PM`.

Using `-StartRow 1 -EndRow 7` imported six data records because the first row contained the headers.

## Creating Worksheets and a Chart

My workflow was:

1. Create the People dataset.
    1. Build four custom objects with Name and City properties.
    2. Display the objects with `Format-Table -AutoSize`.
2. Export the objects to the People worksheet.
    1. Use `-Append -WorksheetName "People"`.
    2. Import the worksheet again to check the saved records.
3. Create the formatted worksheet.
    1. Apply the `Medium16` table style.
    2. Add the title and make it bold with `-TitleBold`.
4. Create the fruit chart.
    1. Use `New-ExcelChartDefinition` with `ColumnClustered`.
    2. Set `Fruits` as the category range and `Quantity` as the value range.
5. Copy the completed workbook to the host and inspect it in WPS Office.

The lab chart example used `$data`, `Item`, and `TotalSold`. I adapted it to the existing `$fruit` variable and the actual column names in the supplied workbook.

## Troubleshooting and Lessons Learned

> A command finishing without an error is only part of the check. I also need to inspect the saved file and its contents.
>
> Checking the worksheet names, imported records, and chart helped me confirm the results.

I also ran a file-copy command in the wrong PowerShell window. The virtual machine's local session did not contain the host's `$session` variable, so `Copy-Item -FromSession` failed.

I returned to the host PowerShell window, exited the interactive remote session, and copied the workbook successfully. This helped me understand the difference between the virtual machine console and a remote session opened from the host.

## Completion Checklist

- [x] Load ImportExcel on the server.
- [x] Import the fruit data and inspect its properties.
- [x] Export process information to Excel.
- [x] Import selected columns and rows.
- [x] Create and verify the People worksheet.
- [x] Create a formatted table.
- [x] Generate and visually check the fruit chart.
- [ ] Finish the submission report with labelled screenshots, explanations, references, and reflection.

## Reflection

Before this lab, I mainly thought of Excel files as documents that needed to be edited manually. I learned how PowerShell objects can become worksheet rows and how a script can create a chart. The most confusing part was switching between the host and virtual machine windows. Checking the prompt and file paths helped me fix this. I also learned to inspect column names instead of assuming that example column numbers always match my file.

## Learning Resources

The course materials used for this work were *WorkingWithExcel1.md*, *Working with Excel*, and *Markdown Language for GIT documentatiion*.

This article is stored in my [PowerShell Excel Lab repository](https://github.com/xlt95260124-sketch/powershell-excel-lab).
