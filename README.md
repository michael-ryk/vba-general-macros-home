# vba-general-macros-home
## What this file contains ?
- Various macros used for different purposes in different excel workbooks
- It run from Personal hidden workbook and applied to any excel used on current PC

## How this file works with excel
- This file only for source control so it is `<file name>.vba` which not really used by excel workbook
- My personal.xlsb excel workbook have module `GeneralMacros` which import-export bas file: `GeneralMacros.bas`
- bas file is like binary for git so i decided to use here just like text files to track changes inside file
- When i write script, i do it inside visual basic editor in excel and not here
- As soon as i finsh and want to save, i should copy it from vba editor to `GeneralMacros.vba` and commit
- Same is in opposite - select all from here and copy all to bas module file

## All Macros
### Emphasize Similar
#### What is used for ?
- When you learn something, it may be hard to keep all the stuff in the head
- This tool created To simplify learning process and reach maximum retain
- Tool applied on Excel data table where each line represent 1 knowledge unit
- Each knowledge unit have following elements: 
  - Subject - Knowledge Question for example "How to add CSS styles to HTML document?" 
  - Answer - You should check this cell after you activly recalled answer after read the question
  - Tags - The most important part. Based on tags you give, connections between notes will be made
  - Date - Last date when you focused on this item
- Tool use popular learning concepts like active recall, spaced repetition, connections between notes
- Main feature of the tool is to show you all connected knoledge units to one you selected

#### How it works ?
- You should have already excel with some data - Each row represent knowledge unit
- Table contain Tag column so for each row you type tags separated by space
- You can put button for easy macro execution
- Now select some row with some knowledge and activate macro
- Macro will do: 
  - Indicate for you previous knowledge unit you focused on
  - All rows which contain at least 1 tag from list of tags on selected row will be markered as "Match" and filtered in the end
  - All rows which contain at least 1 tag from list of tags in subject cell will be markered as "Sugest" to inform you that data may be connected by you using tag
  - After macro finish, it provide you only relevant data that connected to row you focused on 
- To clear, just select some area above table and activate macro - it will releas all filters and apply same status for all rows.
- You may apply this macro to any table like data - You only have to add manuuly required columns as described in install section
- You may lock specific rows if you still want to see over filtered data - they will remain after macro application
- You can see number of connciotns each row have with other notes and by this add to less connected data more connections
- Date will automaticaly update when you focus on topic and this way you con schedule for yourself spaced repetition for topics and ensure you pass over all knowledge items all the time and nothing left behind
- Location column used to provide extra info about where more info can be found - We still have excel here and maybe you have some extra files or video which explains this knowledge topics which can't be added here - so used link.

#### Instalation
- Create new excel worbook with default sheet - Create new data table there starting row 6
- Ensure you table formated as table and contain following headings:
  - Subject
  - Tags
  - Filter
  - Lock
  - Date
  - Connections
  - Found Tag
- In cells F1 and F2 put columns start and end to make range where you want styles to applied (You may want to increase it if your table will have more columns in future)
- You may add slicers you want for more fine tune filters
- Place button and bind it to this macro
- This macro can be as in same excel as in Personal workbook (this way it can be applied on any opened workoob on this pc - read about it if you not familiar with it)
