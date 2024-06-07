# Excel OneDrive Song Linker - an Office Add-in using Microsoft Graph to add Song Links to an Excel spreadsheet

This Excel Add-in uses the Microsoft Graph JavaScript SDK to access folders and files stored in **OneDrive** and insert anonymous read-only links to the selected files starting at the ActiveCell across the current row. If the file is a shortcut then it inserts a link to the target url stored in the shortcut rather than generating a link to the shortcut file. It assumes that there is a base folder containing as set of category folders, and each of the category folders has a set of folders, one per song.  Each song folder contains multiple files with different representations of the song.

By convention it is recommended that the directory names be of the form `[Song Name]_[key]`, and that each file in the directory be of the form `[Song Name]_[key]_[type].[ext]`. The types could include:

- **Lyrics (_lyrics)** - a pdf file with the just the song lyrics.
- **Lead Sheet (_lead)** - a pdf file containing the song melody (and parts) on a staff with the lyrics.
- **Full music (_full)** - a pdf file of full sheet music for the song.
- **Chord chart (_chord)** - a pdf file of the song lyrics and chord progressions written above the words.
- **Individual instrument parts (_<instrument>)** - a pdf file of a particular instrument's part. For example a violin part (_violin).
- **Shortcut to a YouTube performance of the Song (_YouTube)** - a shortcut file containing a link to a YouTube video.

This convention is not required as the Add-in will process all files in the folder and insert them into the spreadsheet. 

# Usage
The Add-in has three selections. 
- **Base Folder** - is a text box where you enter the OneDrive path from the root of your OneDrive to the folder containing the folders of categories of music.  Once this is filled in and you hit **Enter** the Add-in will automatically query OneDrive to get the list of folders in this base folder and populate the **Song Category** drop down list with all the folders contained in the **Base Folder**.  
- **Song Category** - is a drop down list of all the folders containing songs.  Once this drop down is populated from the **Base Folder** the add-in will automatically query OneDrive and populate the **Song** drop down list with all the songs (folders) contained in the **Song Category** folder.  Each time a new **Song Category** is selected, the **Song** drop down list will be changed to have all the songs from the new category.
- **Song** - is a drop down list of all the folders in the **Song Category** folder. 

Once you have selected the **Song** you wish to add the the Excel spreadsheet.  Make sure the **ActiveCell** is pointing to the correct location in the spreadsheet, and then click **Add Song** to perform the insertion. 

The following actions will be taken:
- Query OneDrive for the list of files in the **Base Folder**/**Song Category**/**Song** folder.
- Create a OneDrive **Read Only**-**Anonymous** sharing link to the **Song** folder if it doesn't exist, otherwise use the existing one.
- Set the **ActiveCell** *value* property to the **Song** name, and the *hyperlink* property to a OneDrive **Read Only**-**Anonymous** sharing link to the folder. 
- Iterate through the items in the folder.  Based on the type of item, set the values/properties of sequential cells across the current row.
    - **shortcut (.url)** - read the value of the **shortcut**, set the *value* property of the cell to the filename, and set the *Hyperlink* property of the cell to url referenced by the **shortcut**.
    - **file** (not a shortcut) - create a OneDrive **Read Only**-**Anonymous** sharing link to the **file** if it doesn't exist, otherwise use the existing one, set the *value* property of the cell to the **file** name, and set the *Hyperlink* property of the cell to the **Read Only**-**Anonymous** OneDrive sharing link.
    - **folder** - skipped.

    > [!IMPORTANT]
    > OneDrive **Read Only**-**Anonymous** sharing links allow anyone with the link to be able to read the file.  They will not have access to make any changes to the file.

# Support
This code is open source and the repository is on github: [ExcelOneDriveSongLinker](YOUR_GITHUB_LOCATION).

If you encounter an issue, you can submit and track issues on github: [Issue](YOUR_GITHUB_LOCATION/issues).

