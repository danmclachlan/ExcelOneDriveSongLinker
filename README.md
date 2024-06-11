# Office Add-in using Microsoft Graph to add Song Links to an Excel spreadsheet

![License.](https://img.shields.io/badge/license-MIT-green.svg)

This code is based on the [Microsoft Graph sample Office Add-in](https://learn.microsoft.com/en-us/samples/microsoftgraph/msgraph-sample-office-addin/microsoft-graph-sample-office-add-in/) from the github repository: [msgraph-sample-office-addin](https://github.com/microsoftgraph/msgraph-sample-office-addin/tree/main/).  

This Excel Add-in uses the Microsoft Graph JavaScript SDK to access folders and files stored in **OneDrive** and insert anonymous read-only links to the selected files starting at the ActiveCell across the current row. If the file is a shortcut then it inserts a link to the target url stored in the shortcut rather than generating a link to the shortcut file. It assumes that there is a base folder containing as set of category folders, and each of the category folders has a set of folders, one per song.  Each song folder contains multiple files with different representations of the song.

By convention it is recommended that the directory names be of the form `[Song Name]_[key]`, and that each file in the directory be of the form `[Song Name]_[key]_[type].[ext]`. The types would include:

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

# Instantiate the project
## Prerequisites

To build and run this project, you need the following:

- [Node.js](https://nodejs.org) and [Yarn](https://yarnpkg.com/) installed on your development machine. (**Note:** This was written with Node version 20.13.0 and Yarn version 1.22.22. Other versions have not been tested.

- Either a personal Microsoft account with OneDrive, or a Microsoft work or school account.

If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.

## Register a web application in the Azure portal

1. Open a browser and navigate to the [Azure Portal](https://portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **App registrations** under the **Azure Services** section of the Home page.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Excel OneDrive Song Linker (localhost)`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.

1. Select **Register**. On the **Excel OneDrive Song Linker (localhost)** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the next step.

    > [!IMPORTANT]
    > This client secret is never shown again, so make sure you copy it now.

1. Select **API permissions** under **Manage**, then select **Add a permission**.

1. Select **Microsoft Graph**, then **Delegated permissions**.

1. Select the following permissions, then select **Add permissions**.

    - **Files.ReadWrite.All** - this will allow the app to read and write to the user's OneDrive Files. This is needed to create the read-only anonymous links to the files.
    - **openid** - to sign users in.
    - **profile** - View user's basic profile.
    - **User.Read** - Sign in and read user profile.

## Configure Office Add-in single sign-on

Update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

1. Select **Expose an API**. In the **Scopes defined by this API** section, select **Add a scope**. When prompted to set an **Application ID URI**, set the value to `api://localhost:3000/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID. Choose **Save and continue**. 
    > [!IMPORTANT]
    > This field is pre-populated with your APP_ID, but does NOT include the `localhost:3000`.  It is **VERY** important to add this!

1. Fill in the fields as follows and select **Add scope**.

    - **Scope name:** `access_as_user`
    - **Who can consent?:** `Admins and users`
    - **Admin consent display name:** `Access the app as the user`
    - **Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`
    - **User consent display name:** `Access the app as you`
    - **User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`
    - **State:** `Enabled`

1. In the **Authorized client applications** section, select **Add a client application**. Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**. Repeat this process for each of the client IDs in the list.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
    - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)

## Install development certificates

1. Run the following command to generate and install development certificates for your add-in.

    ```Shell
    npx office-addin-dev-certs install
    ```

    If prompted for confirmation, confirm the actions. Once the command completes, you will see output similar to the following.

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. Copy the paths to localhost.crt and localhost.key, you'll need them in the next step.

## Configure the code
The code is structured to be able to run both in a local development environment using https://localhost:3000 and to be deployed to Azure using App Services (https://AppName.azurewebsites.net).
It is highly recommended to create a second **App Registration** for the deployed version (see below for details).

1. Rename `template.privateEnvOptions.js` to `privateEnvOptions.js`
1. Edit the `privateEnvOptions.js` file and make the following changes to the `envLocal` object.
    1. Replace `<your new local guid>` with a new guid. You can use the `New-Guid` function in PowerShell to generate a new guid.
    1. Replace `<your localhost CLIENTID>` with the **Application Id** you got from the App Registration Portal.
    1. Replace `<your localhost Client Secret>` with the client secret you got from the App Registration Portal.
    1. Replace `<path to LOCALHOST.CRT>` with the path to your localhost.crt file from the output of the `npx office-addin-dev-certs install` command. You will need to replace all \ with a double \ in the path.
    1. Replace `<path to LOCALHOST.KEY>` with the path to your localhost.key file from the output of the `npx office-addin-dev-certs install` command. You will need to replace all \ with a double \ in the path.
    1. Save the file.

1. In your command-line interface (CLI), navigate to this directory and run the following command to install requirements.

    ```Shell
    yarn install
    ```

## Build and run the code

1. Run the following command in your CLI to build the application. The built code will be placed in the `dist/localhost` folder.

    ```Shell
    yarn build:devpack
    ```
1. Run the following command in your CLI to run the application.

    ```Shell
    yarn start:devpack
    ```

1. In your browser, go to [Office.com](https://www.office.com/) and sign in. Select **Create** in the left-hand toolbar, then select **Workbook**.

1. Select the **Insert** tab, then select **Office Add-ins**.

1. Select **Upload My Add-in**, then select **Browse**. Upload your **manifest.xml** file. It will be in the `dist/localhost` folder.

1. Select the **Load Song** button on the **Home** tab to open the task pane.


## Deploying to Azure App Services
This process follows parts of the process outlined in the Microsoft Learn documentation [Deploy a single sign-on (DDO) Office Add-in to Microsoft Azure App Service](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/deploy-office-add-in-sso-to-azure). Individual steps to will be highlighted but not duplicated here.

1. Additional Requirements
    - An Azure account. Get a trial subscription at [Microsoft Azure](https://azure.microsoft.com/free/).
    - [Azure App Service extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) for VS Code.
    
1. [Create the App Service App](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/deploy-office-add-in-sso-to-azure?tabs=windows#create-the-app-service-app) - Follow the steps in this section,  stoping when you get to the **Update package.json** section. 
    - It is important to make sure to set the `SCM_DO_BUILD_DURING_DEPLOYMENT` to `true` in the `Application Settings`.
    - Make sure to capture domain name **URL** (not the `https://` part) from the Overview pane in the Azure Portal. You'll need itin later steps.

1. Create a new App Registration
    1. Open a browser and navigate to the [Azure Portal](https://portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

    1. Select **App registrations** under the **Azure Services** section of the Home page.

    1. Select **New registration**. On the **Register an application** page, set the values as follows.

        - Set **Name** to `Excel OneDrive Song Linker (Azure)`.
        - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
        - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://<App Service app URL saved above>/consent.html`.

    1. Select **Register**. On the **Excel OneDrive Song Linker (Azure)** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

    1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

    1. Copy the client secret value before you leave this page. You will need it in the next step.

        > [!IMPORTANT]
        > This client secret is never shown again, so make sure you copy it now.

    1. Select **API permissions** under **Manage**, then select **Add a permission**.

    1. Select **Microsoft Graph**, then **Delegated permissions**.

    1. Select the following permissions, then select **Add permissions**.

        - **Files.ReadWrite.All** - this will allow the app to read and write to the user's OneDrive Files. This is needed to create the read-only anonymous links to the files.
        - **openid** - to sign users in.
        - **profile** - View user's basic profile.
        - **User.Read** - Sign in and read user profile.

1. Configure Office Add-in single sign-on

    Update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

    1. Select **Expose an API**. In the **Scopes defined by this API** section, select **Add a scope**. When prompted to set an **Application ID URI**, set the value to `api://<App Service app URL saved above>/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID. Choose **Save and continue**. 
    1. Fill in the fields as follows and select **Add scope**.
    - **Scope name:** `access_as_user`
    - **Who can consent?:** `Admins and users`
    - **Admin consent display name:** `Access the app as the user`
    - **Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`
    - **User consent display name:** `Access the app as you`
    - **User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`
    - **State:** `Enabled`

    1. In the **Authorized client applications** section, select **Add a client application**. Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**. Repeat this process for each of the client IDs in the list.

        - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
        - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
        - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
        - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)

1. Create a new App registration. Follow the steps outlined above in **Register a web application in the Azure portal** with the following differences:
    1. Set **Name** to `Excel OneDrive Song Linker (Azure)`.
    1. Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://<App Service app URL saved above>/consent.html`.
    1. In the Select **Expose an API**. In the **Scopes defined by this API** section, select **Add a scope**. When prompted to set an **Application ID URI**, set the value to `api://<App Service app URL saved above>/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID. Choose **Save and continue**. 

1. Edit the `privateEnvOptions.js` file and make the following changes to the `envAzure` object.  
    1. Replace `<your new Azure guid>` with a new guid. You can use the `New-Guid` function in PowerShell.
    1. Replace `<your Azure CLIENTID>` with the **Application Id** you got from the App Registration Portal.
    1. Replace `<your Azure Client Secret>` with the client secret you got from the App Registration Portal.
    1. Replace `<App Service app URL>` with the URL for your app domain.  It wil be of the form `<name>.azurewebsites.net`.
    1. Save the file.

1. Build and deploy the code - follow the steps in [Build and deploy](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/deploy-office-add-in-sso-to-azure?tabs=windows#build-and-deploy) except use the following command to build the code (the built code will be in the `dist/azure` folder):

    ```Shell
    yarn build
    ```

1. In your browser, go to [Office.com](https://www.office.com/) and sign in. Select **Create** in the left-hand toolbar, then select **Workbook**.

1. Select the **Insert** tab, then select **Office Add-ins**.

1. Select **Upload My Add-in**, then select **Browse**. Upload your **manifest.xml** file. It will be in the `dist/azure` folder.

1. Select the **Load Song** button on the **Home** tab to open the task pane.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
