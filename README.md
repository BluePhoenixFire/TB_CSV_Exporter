# TB_CSV_Exporter
A simple macro to check TB's balance then export to CSV for re-upload into consolidation systems

**Overview**
This repository contains a VBA (Visual Basic for Applications) macro designed for use in Microsoft Excel. The macro automates the process of checking if the sum of a specified range in a worksheet equals zero. If the sum is within a defined tolerance level of zero, the macro exports the worksheet to a CSV file.

**Features**
Sum Check: Calculates the sum of a user-defined range in a worksheet.
Tolerance Level: Uses a defined tolerance level to account for floating-point arithmetic limitations.
CSV Export: If the sum is within the tolerance of zero, the worksheet is exported as a CSV file to a specified location.

**Prerequisites**
Microsoft Excel
Basic understanding of Excel and VBA

**Installation**
Open the Excel workbook where you want to use the macro.
Press Alt + F11 to open the VBA editor.
Import the macro file or copy and paste the macro code into a new module.

**Usage**
Adjust the sumRange variable in the macro to reflect the range you want to sum.
Set the tolerance variable to your desired level of precision.
Ensure the file name for the CSV export is specified in the designated cell (default is J1).
Run the macro to perform the sum check and export if conditions are met.

**Customization**
Modify the sumRange to change the range of cells being summed.
Adjust the tolerance level as per the precision requirements of your task.
Change the cell reference for fileName to use a different cell for naming the exported CSV file.

**Limitations**
The macro currently does not support dynamic range selection or automatic tolerance adjustment.

**Author**
BluePhoenix
