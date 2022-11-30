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