# vba-general-macros-home
## Where this scripts used
This file contains various macros used in various excels on my home pc

## How this file works with excel
- This file only for source control so it is `<file name>.vba` which not really used by excel workbook
- My personal.xlsb excel workbook have module `GeneralMacros` which import-export bas file: `GeneralMacros.bas`
- bas file is like binary for git so i decided to use here just like text files to track changes inside file
- When i write script, i do it inside visual basic editor in excel and not here
- As soon as i finsh and want to save, i should copy it from vba editor to `GeneralMacros.vba` and commit
- Same is in opposite - select all from here and copy all to bas module file

## All Macros
### Emphasize Similar
#### Wha is used for ?
- Used to focus on knowledge topics listed in excel table by show only rows which have at least one of the tags words from taps column of focused row
- It tracks previously focused row and mark it blue on next run
- All rows which have at least 1 tag from focused rows emphesized by bold
- All rows which have at least 1 tag word found in subject, marked as light grey to show uncertain connections
- All remain rows colored very light grey to reduce focus
- All rows sorted by importance

#### What need to run this macro ?
- You should have Data table which Start at least from 3th row and contains at least following column names:
  - Subject
  - Tags
  - Filter
  - Lock
  - Date
  - Connections

#### How it works ?
- Macro detects table automaticaly on active sheet and search for columns above to determine their address
- Next, it just run and do the work