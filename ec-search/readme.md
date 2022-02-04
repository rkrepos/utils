# Excel Custom Search
Search macro that hides non-matching rows. Once installed, the macro is loaded in the background everytime Excel is opened, allowing it to be used on any open workbook. 

### Installation
* Copy **PERSONAL.XLSB** to location **`C:\Users\<UserName>\AppData\Roaming\Microsoft\Excel\XLSTART`**
* If you already have a **PERSONAL.XLSB** in the above location
	* Rename the new **PERSONAL.XLSB** to allow opening it.
	* Open both the files and copy or merge all the code from the new one to existing one as needed.

### How To Use
* To Search
	* Press **`Ctrl + Shift + H`** to bring up the Search box.
	* Type in the value to search for.
	* Press **`Enter`** key or click the **`Ok`** button.
* To Unhide All
	* Press **`Ctrl + Shift + J`**

### How To Modify
* Change ShortCut Keys
	* Open the **PERSONAL.XLSB** and go to **`Developer > Visual Basic > VBAProject (PERSONAL.XLSB) > Microsoft Excel Objects > Workbook`**
	* Modify the ShortCut strings provided on top.
* Change Search Code
	* Open the **PERSONAL.XLSB** and go to **`Developer > Visual Basic > VBAProject (PERSONAL.XLSB) > Modules > Search`**
	* Modify the code as needed.
