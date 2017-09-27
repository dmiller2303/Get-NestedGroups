# Get-NestedGroups

Gets the nested GROUPS (only groups) of the specified top level group.  By using the `WithExcel` parameter, you can output the results to an excel file.  The console results, as well as the excel file, will be in a tree style colored output.

Example:
```
  Top Level Group
    NestedGroup1
      NestedGroup1.1
        NestedGroup1.1.1
      NestedGroup1.2
    NestedGroup2
    NestedGroup3
```
Parameters:

- ToplevelGroup - The name of the Top Level Group to be searched
- WithExcel - Switch Parameters to indicate if you want to output to Excel or just to the Console
- ExcelFilePath - If outputting to Excel, the desired Excel file
