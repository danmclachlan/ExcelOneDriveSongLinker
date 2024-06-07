# Excel OneDrive Song Linker - Installation Guide

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



