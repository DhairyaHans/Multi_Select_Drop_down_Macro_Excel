# ExcelMacros

## Files ->
  
  # Multi_select_drop_down_in_excel -> 
    - The file serves the purpose of auto updating the lists and its values.
    - Say, We have 2 columns, 
      * A -> having comma separated values which act as source for our drop down list column.
      * B -> it has the drop-down list, for corresponding row in the A column
    - This file, performs mainly 2 functions ->
      * Update the drop-down list options of corresponding cell in the column B; for any change in the comma separated values of a cell in the column A.
      * Update the drop-down list values of the cells, based on selection of values from a cell in the column B.
    - Each drop-down list has an option of 'Clear' which removes all the values present in the corresponding cell.
    - It handles the case, where We can add new values in the column A and it automatically updates the corresponding column in the drop-down column B.

  # split_values_to_list ->
    - This file contains the helper function, SplitValuesToList, which splits the string (having comma-separated values) passed as inputString and converts it 
        to a list of values; with an additional value of 'Clear'

  # update_drop_down_lists ->
    - This file works on the Existing values present in the Excel sheet.
    - While using Excel, We may have the case, where, We have data already present, and we need to add the functionality of multi-select dropdown to the sheet.
    - Running this macro file, updates/creates the drop-down list in the column B, based on the existing values present in the cells of column A.

# IMPORTANT NOTE -> 
  - **Multi_select_drop_down_in_excel** file, is important for handling the updates that will be happening in the file columns.
  - **update_drop_down_lists** file,is important for handling the drop-down list for existing values in the columns.
  

## Steps to create a multi select drop down list without repeatition in a column, based on comma separated values present in another column ->
    eg -> 
    
    Col A          Col B
    X,Y,Z          (Drop down list having values -> X, Y, Z, and Clear as list options to select from)
    P,Q,R,S        P, Q  (Assuming P and Q are selected)

    Steps ->
        Step 1: Press Alt + F11 to open the VBA Editor.
        Step 2: Go to Insert > Module to insert a new module.
        Step 3: Copy and paste the code of the "**split_values_to_list**" file into the module window.
        Step 4: Go to Insert > Module to insert a new module.
        Step 5: Copy and paste the code of the "**update_drop_down_lists**" file into the module window.
        Step 6: In the Project Explorer window, find your worksheet name under "Microsoft Excel Objects" (e.g., "Sheet12 (Hello World)").
        Step 7: Double-click on the worksheet name to open the code window for that worksheet.
        Step 8: Under the (General) tab, select "Worksheet" and under the (Declarations) tab, select "Change".
        Step 9: Copy and paste the code of the "**Multi_select_drop_down_in_excel**" file, into the window.
        Step 10: Run the Macro, update_drop_down_lists ("UpdateDropDownLists"), by pressing Alt + F8 and selecting the macro.
        Step 11: Now, you can use the functionality of multi-select drop down on column A (Working as Source) and Column B (Working as Target)
  
# IMPORTANT POINTS ->
  - In the code, I have used column 'F' as my Source and column 'G' as my Target Column.
  - Also, Do update the Sheet name, based on your sheet, I have used "Sheet12".
  - You can find your sheet name in the VBA editor, in the Project Explorer window, (say, "Sheet1 (World)"), then you have to use "Sheet1" as your sheet name.
