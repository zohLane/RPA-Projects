# RPA-Projects
Projects of Custom Activities

Pivot of Sum uses CodeActivity and helps to create Pivot in Excel for:
- A Row Field
- A Value Field
- A Sum applied to the Value Field

1. The data is aggregated by PivotRowField  (Required Argument)

2. Sum is calculated on PivotValueField (Required Argument)

3. SourceSheet and SourceRange specifies the data for pivot (Required Argument)
    E.g:
      SourceSheet = "Sheet1"
      SourceRange = "A1:Z100"

4. TargetSheet, TargetRowIndex, TargetColumnIndex specify the location of pivot (Required Argument). TargetSheet has to be in the same file and will be created, if it does not exist.
     E.g:
        TargetRowIndex = 1 (To specify excel row 1)
        TargetColumnIndex = 1 (To specify excel column A)

You cannot write the Pivot in the same location as another pivot.

5. TableName is an optional Argument to use for naming the table (Default value - Pivot of Sum).You cannot have the same Table Name for multiple pivot tables in the same excel sheet.

6. Close the excel file before calling on this activity
