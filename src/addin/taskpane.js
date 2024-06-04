// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <AuthUiSnippet>
// Handle to authentication pop dialog
/**
 * @type {Office.Dialog | undefined}
 */
let authDialog = undefined;

let OfficeRT = undefined;

/**
 * Retrieves an access token for the Office Add-in's web application.
 * @param {OfficeRuntime.AuthOptions} [options] - Optional authentication options.
 * @returns {Promise<string>} - A promise that resolves with the access token.
 */
async function getAccessToken(options) {
  if (OfficeRT === undefined) {
    console.debug('getAccessToken: OfficeRT not defined');
    // @ts-ignore
    // eslint-disable-next-line @typescript-eslint/no-unused-vars, no-undef
    OfficeRT = OfficeRuntime || { auth: { getAccessToken: () => {throw new DOMException('getAccessToken: office.js not loaded', 'NotFoundError');}}};
  }
  return await OfficeRT.auth.getAccessToken(options);
}

// Build a base URL from the current location
function getBaseUrl() {
  return location.protocol + '//' + location.hostname +
  (location.port ? ':' + location.port : '');
}

// Process the response back from the auth dialog
/**
 * @param {{ message: string; origin: string | undefined; } | { error: number }} result
 */
function processConsent(result) {
  // @ts-ignore
  const message = JSON.parse(result.message);

  authDialog?.close();
  if (message.status === 'success') {
    showMainUi();
  } else {
    const error = JSON.stringify(message.result, Object.getOwnPropertyNames(message.result));
    showStatus(`An error was returned from the consent dialog: ${error}`, true);
  }
}

// Use the Office Dialog API to show the interactive
// login UI
function showConsentPopup() {
  const authDialogUrl = `${getBaseUrl()}/consent.html`;

  Office.context.ui.displayDialogAsync(authDialogUrl,
    {
      height: 60,
      width: 30,
      promptBeforeOpen: false
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        authDialog = result.value;
        authDialog.addEventHandler(Office.EventType.DialogMessageReceived, processConsent);
      } else {
        // Display error
        const error = JSON.stringify(result.error, Object.getOwnPropertyNames(result.error));
        showStatus(`Could not open consent prompt dialog: ${error}`, true);
      }
    });
}

// Inform the user we need to get their consent
function showConsentUi() {
  $('.container').empty();
  $('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: 'Consent for Microsoft Graph access needed'
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'In order to access your files, we need to get your permission to access the Microsoft Graph.'
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'We only need to do this once, unless you revoke your permission.'
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'Please click or tap the button below to give permission (opens a popup window).'
  }).appendTo('.container');
  $('<button/>', {
    class: 'primary-button',
    text: 'Give permission'
  }).on('click', showConsentPopup)
    .appendTo('.container');
}

// Display a status
/**
 * @param {unknown} message
 * @param {boolean} isError
 */
function showStatus(message, isError) {
  $('.status').empty();
  $('<div/>', {
    class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
  }).append($('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: isError ? 'An error occurred' : 'Success'
  })).append($('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: message
  })).appendTo('.status');
}

/**
 * @param {boolean} show
 */
function toggleOverlay(show) {
  $('.overlay').css('display', show ? 'block' : 'none');
}
// </AuthUiSnippet>

// <MainUiSnippet>
function showMainUi() {
  $('.container').empty();

  $('<hr2/>', {
    class: 'ms-fontSize-24 ms-fontWeight-semibold',
    text: 'Select song to add'
  }).appendTo('.container');
  $('<hr2/>', {
    class: 'ms-fontSize-12 ms-fontWeight-semibold',
    text: ' (at ActiveCell)'
  }).appendTo('.container');

  // Create the basefolder form
  $('<form/>').on('change', getSongCategories)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Base Folder'
    })).append($('<input/>', {
      class: 'form-input',
      type: 'text',
      required: true,
      id: 'baseFolder'
    })).appendTo('.container');

  // Create the Song Category input form
  $('<form/>').on('change', getSongOptions)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Song Category'
    })).append($('<select/>', {
      id: 'categoryFolder',
      class: 'form-input',
      type: 'text',
      required: true,
    })).appendTo('.container');
  
  // Create the input form
  $('<form/>').on('submit', getSongLinks)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Song'
    })).append($('<select/>', {
      id: 'songName',
      class: 'form-input',
      type: 'text',
      required: true,
    })).append($('<input/>', {
      class: 'primary-button',
      type: 'submit',
      id: 'importButton',
      value: 'Add Song'
    })).appendTo('.container');

  $('<hr/>').appendTo('.container');
}
// </MainUiSnippet>

// <getSongCategoriesSnippet>
/**
 *  @param {{ preventDefault: () => void; }} evt
 */
async function getSongCategories(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  console.debug('getSongCategories ...');

  try {
    const apiToken = await getAccessToken({ allowSignInPrompt: true });
    const baseFolder = $('#baseFolder').val();
    const requestUrl =
      `${getBaseUrl()}/graph/folderchildren?baseFolder=${baseFolder}`;

    const response = await fetch(requestUrl, {
      headers: {
        authorization: `Bearer ${apiToken}`
      }
    });

    if (response.ok) {
      const categoryList = await response.json();
      let categoryCount = 0;
      if (categoryList.length > 0) {
        $('#categoryFolder').empty();

        const selectElement = document.getElementById('categoryFolder');
        categoryList.forEach((/** @type {{ Name: string; Type: string; ChildCount: Number | null; }} */ element) => {
          if (element.Type === 'Folder') {
            const option = document.createElement('option');
            option.value = element.Name;
            option.textContent = element.Name;
            selectElement?.appendChild(option);
            categoryCount++;
          }
        });
      }

      showStatus(`got ${categoryCount} Song Categories`, false);
      
      // Fill in the Songs from the first category.
      getSongOptions(evt);

    } else {
      const error = await response.json();
      showStatus(`Error populating Song Category List from OneDrive: ${JSON.stringify(error)}`, true);
    }
    
  } catch (err) {
    console.log(`Error: ${JSON.stringify(err)}`);
    showStatus(`Exception populating Song Category List from OneDrive: ${JSON.stringify(err)}`, true);
  }

  toggleOverlay(false);
}
// </getSongCategoriesSnippet>

// <getSongOptionsSnippet>
/**
 *  @param {{ preventDefault: () => void; }} evt
 */
async function getSongOptions(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  console.debug('getSongOptions ...');

  try {
    const apiToken = await getAccessToken({ allowSignInPrompt: true });
 
    let baseFolder = $('#baseFolder').val()?.toString() || '';
    const categoryFolder = $('#categoryFolder').val()?.toString() || '';
    //console.debug('categoryFolder: ', categoryFolder);

    if (categoryFolder !== '') {
      baseFolder = baseFolder.concat('/', categoryFolder);
      //console.debug('updated baseFolder: ', baseFolder);
    }

    const requestUrl =
      `${getBaseUrl()}/graph/folderchildren?baseFolder=${baseFolder}`;

    const response = await fetch(requestUrl, {
      headers: {
        authorization: `Bearer ${apiToken}`
      }
    });

    if (response.ok) {
      const songList = await response.json();
      if (songList.length > 0) {
        $('#songName').empty();

        const selectElement = document.getElementById('songName');
        songList.forEach((/** @type {{ Name: string; Type: string; ChildCount: Number | null; }} */ element) => {
          if (element.Type === 'Folder') {
            const option = document.createElement('option');
            option.value = categoryFolder.concat('/', element.Name);
            option.textContent = element.Name;
            selectElement?.appendChild(option);
          }
        });
      }

      showStatus(`got ${songList.length} songs`, false);
    } else {
      const error = await response.json();
      showStatus(`Error populating Song List from OneDrive: ${JSON.stringify(error)}`, true);
    }
    
  } catch (err) {
    console.log(`Error: ${JSON.stringify(err)}`);
    showStatus(`Exception populating Song List from OneDrive: ${JSON.stringify(err)}`, true);
  }

  toggleOverlay(false);
}
// </getSongOptionsSnippet>

// <getSongLinksSnippet>
/**
 * @param {{ preventDefault: () => void; }} evt
 */
async function getSongLinks(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  try {
    const apiToken = await getAccessToken({ allowSignInPrompt: true });
 
    const baseFolder = $('#baseFolder').val();
    const songName = $('#songName').val();

    const requestUrl =
      `${getBaseUrl()}/graph/folderitemsurls?baseFolder=${baseFolder}&songName=${songName}`;

    const response = await fetch(requestUrl, {
      headers: {
        authorization: `Bearer ${apiToken}`
      }
    });

    if (response.ok) {
      const files = await response.json();
      if (files.length > 0) WriteUrlsToSheet(files);
      showStatus(`Inserted ${files.length} song file links`, false);
    } else {
      const error = await response.json();
      showStatus(`Error getting links from OneDrive: ${JSON.stringify(error)}`, true);
    }

    
  } catch (err) {
    console.log(`Error: ${JSON.stringify(err)}`);
    showStatus(`Exception getting links from OneDrive: ${JSON.stringify(err)}`, true);
  }
  toggleOverlay(false);
}
// <!<getSongLinksSnippet>

// <WriteUrlsToSheetSnippet>
/**
 * @param {any[]} items
 */
async function WriteUrlsToSheet(items) 
{
  console.debug(`in WriteUrlsToSheet: items count = ${items.length}`);
  await Excel.run(async (context) => 
  {
    const cell = context.workbook.getActiveCell();
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject();

    cell.load('address');
    cell.load('rowIndex');
    cell.load('columnIndex');
    usedRange.load('isNullObject');
    usedRange.load('columnCount');

    await context.sync();

    //console.debug('Used Range: ', usedRange);

    if (!usedRange.isNullObject) {
      // the sheet is not blank so make sure the row we are inserting
      // into is empty starting at the ActiveCell 
      // get the range to clear 
      let rangeToClear = sheet.getRangeByIndexes(cell.rowIndex, cell.columnIndex + 1, 1, usedRange.columnCount - cell.columnIndex + 1);

      //console.debug('Range to clear: ', rangeToClear);
      rangeToClear.clear();
      await context.sync();
    }

    //get the full address of the active cell
    var activeCellAddress = cell.address;

    // calculate the range needed to insert the hyperlinks
    var range = sheet.getRange(activeCellAddress).getResizedRange(0, items.length - 1);
    range.load('rowCount');
    range.load('columnCount');
    range.load('cellCount');
    await context.sync();

    //console.debug('Range to insert: ', range);

    for (var i = 0; i < items.length; i++) {
      let cell = range.getCell(0,i);
      cell.hyperlink = {
        address: items[i].WebUrl,
        textToDisplay: items[i].Name
      };
      cell.values = items[i].Name;

      range.format.autofitColumns();
    }
  }).catch (function(err) {
    console.log(`Error: ${JSON.stringify(err)}`);
    showStatus(err, true);
  });
}
// </WriteUrlsToSheetSnippet>

// <OfficeReadySnippet>
Office.onReady(info => {
  console.debug('in Office.onReady');
  // Only run if we're inside Excel
  if (info.host === Office.HostType.Excel) {
    // eslint-disable-next-line no-undef
    OfficeExtension.config.extendedErrorLogging = true; // better debugging
    $(async function() {
      let apiToken = '';
      try {
        apiToken = await getAccessToken({ allowSignInPrompt: true });
        //console.debug('Office.onReady: API Token: ', apiToken);
      } catch (error) {
        console.debug(`Office.onReady: getAccessToken error: ${JSON.stringify(error)}`);
        // Fall back to interactive login
        showConsentUi();
      }

      // Call auth status API to see if we need to get consent
      const authStatusResponse = await fetch(`${getBaseUrl()}/auth/status`, {
        headers: {
          authorization: `Bearer ${apiToken}`
        }
      });

      const authStatus = await authStatusResponse.json();

      console.debug(`auth/status response: ${JSON.stringify(authStatus)}`);

      if (authStatus.status === 'consent_required') {
        showConsentUi();
      } else {
        // report error
        if (authStatus.status === 'error') {
          const error = JSON.stringify(authStatus.error,
            Object.getOwnPropertyNames(authStatus.error));
          showStatus(`Error checking auth status: ${error}`, true);
        } else {
          showMainUi();
        }
      }
    });
  }
});
// </OfficeReadySnippet>