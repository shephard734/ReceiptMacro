# ReceiptMacro
An Excel VBA macro that will help my coworker Kristen save time.

Currently, she spends hours every week manually copying notes from last week's order receipts spreadsheet.
This macro will cut that time down drastically by allowing her to select a previous week's spreadsheet and import them automatically, as long as she has already set both the source and destination sheets up the way she wants them.

The structure goes like this:
- Request that an Excel workbook be chosen with a file browser dialog box and open it.
- Copy the first sheet in that workbook to the destination workbook (whichever one was open when the macro was run).
- Name a range that includes each order's confirmation and notes data as well as its unique order number.
- Set a formula on each destination cell that uses vlookup to match its order number with that in the named range.
- Convert the results of that formula to static values.
- Delete the copied sheet and close the source workbook.

Future Versions, if there are any, might include:
- Copying over the color formatting of imported data or applying conditional formatting that will render that unnecessary.
- Automatically transforming the spreadsheet so that Kristen doesn't have to do so manually before running this macro.
