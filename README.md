# Endless Sky ships and outfits data importer for Excel

Macros that fetch ships and outfits data from Endless Sky folder to an Excel file.

**Only tested on Windows 10**


## Installation guide for those who are not familiar with VBA

1. Download all the .bas files in Module folder of this repository

2. Create a blank Excel file that enabled macros (.xlsm)

3. Open the VBA editor through Developer > Visual Basic

4. In the VBA editor, click File > Import File... 

5. Select a .bas file, click Open

6. Repeat step 4-5 for every other .bas files	;)

7. If your game folder is NOT at C:\Program Files (x86)\Steam\steamapps\common\Endless Sky

	1. Look at the project menu on the left, find SheetSetup under Modules

	2. find the variable named gameFilePath

	3. edit it so that `gameFilePath = "<your game folder path>"`

8. Go back to the Excel, click Developer > Macros, run Setup

9. Re-adjust cell sizes to make them readable

10. Re-run the macros for specific sheets (e.g. Ship.Data for ships) if you have made changes to their headings


## Customize the (limited) attributes you want to fetch 

- You can edit the content and number of headings to fetch the data you want, go into the game file listed in filepath to find the valid attributes

- Invalid attributes or typo will give you no data for that column

- headings are not case-sensitive

- the macros will start scanning for headings from cell C2 to the right, and stop once it scanned an empty cell or cell with only space

- Not intened for fetching visual attributes like sprites, though they may work

- Cannot fetch dafault outfits for ships (http://endless-sky.7vn.io/ should cover all your needs in this aspect ~~and all other aspects, actually this whole project is redundant~~)

## Update-It-Yourself

- New species? Add their data file paths to the "filepath" sheet, ships on the left, outfits on the right (and pray that the macros can read them)

- Balance change? Just re-run the macros

- the macros will start scanning for paths from cell B3(for ships)/D3(for outfits) to the bottom, and stop once it scanned an empty cell or cell with only space

## Read data from plugins

Errrrrrr this should work? I suppose? As long as it follows the data file format guidelines, I think. I never looked into how the game read the data files so the macros probably don't read the files like the game does. So it's likely that on some occasions, the game can read the data while the macros cannot.