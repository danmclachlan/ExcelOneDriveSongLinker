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
This code is open source and the repository is on github [ExcelOneDriveSongLinker](https://github.com/danmclachlan/ExcelOneDriveSongLinker).

If you encounter an issue, you can submit and [Issue](https://github.com/danmclachlan/ExcelOneDriveSongLinker/issues).

# Installation
To use the **Excel OneDrive Song Linker** add-in, you will need to use sideloading to install this Office Add-in. There are different ways to sideload this Office Add-in based on the specific platform.  The sideloading process is detailed in Microsoft documentation - [Sideload an Office Add-in for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).  Two of the common ones for Windows are detailed below.
- **Sideload in Office on the web**
- **Sideload on Windows from a network share**

In all cases you will need the `manifest.xml` file for this Office Add-in. This manifest file is available on the webserver that implements the backend for this Add-in and is all that is needed to uniquely identify this Add-in. It is located at: `https://localhost:3000/manifest.xml`.

## Retrieve the `manifest.xml` file
Execute the following steps to get the `manifest.xml` file for the **Excel OneDrive Song Linker** Office Add-in.

1. Open a browser and go to the site `https://localhost:3000/manifest.xml`.
1. Right click in the browser window and select **Save as**.
1. Pick a location to save the file.

## Sideload in Office on the web
1. Open [Office on the web](https://office.live.com/) and open a **Excel** document. 
1. Select **Home >> Add-ins**, and then select **More Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
1. **Browse** to the folder where you saved the `manifest.xml` file in the **Retrieve the `manifest.xml` file** above, and then select **Upload**.
1. You should see a pop-up notification that it is installed.
1. You shoulld now see a **Load Song** button in the **Home** tab of the ribbon. Click on it to load the add-in. 

## Sideload on Windows from a network share
The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.
[Sideloading Office Add-ins into Office Desktop](https://youtu.be/XXsAw2UUiQo).

Here are the steps:
1. Share a folder
1. Specify the shared folder as a trusted catelog
1. Side load the Add-in

### Share a folder
1. In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder where you saved the `manifest.xml`.  This will become your shared folder catalog.
1. Open the context menu for the folder you want to use as your shared folder catalog (for example, right-click the folder) and choose **Properties**.
1. Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.
1. Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share this add-in. 
You'll need at least **Read/Write** permission to the folder. After you've finished choosing people to share with, choose the **Share** button.
1. When you see the **Your folder is shared** confirmation, make note of the full network path that's displayed immediately following the folder name. 
(You'll need to enter this value as the **Catalog Url** when you specify the shared folder as a trusted catalog, as described in the next section.) 
Choose the **Done** button to close the **Network access** dialog window.
1. Choose the **Close** button to close the Properties dialog window.

### Specify the shared folder as a trusted catelog
1. Open a new document in Excel.
1. Choose the **File** tab, and then choose **Options"".
1. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
1. Choose **Trusted Add-in-Catelogs**.
1. In the **Catelog Url** box, enter the full network path to the folder that you shared previously. If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window and the **Sharing** tab, and look for **Network Path:**.
1. After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.
1. Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.
1. Choose the **OK** button to close the **Options** dialog window.
1. Close and reopen the **Excel** application so your changes will take effect.

### Side load the add-in
1. in Excel, select **Home > Add-ins** from the ribbon, then select **Get Add-ins**.  
1. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.
1. Select the name of the add-in and choose **Add** to insert the add-in.
1. You should see a pop-up notification that it is installed.
1. You shoulld now see a **Load Song** button in the **Home** tab of the ribbon. Click on it to load the add-in. 

## Removing a sideloaded add-in
You can remove a previously sideloaded add-in by clearing the Office cache on your computer. The details can be found in 
[Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).



