// Copyright (c) Microsoft Corporation.
// Copyright (c) Daniel R. McLachlan.
// Licensed under the MIT License.

// <AuthUiSnippet>
// Handle to authentication pop dialog
/**
 * @type {Office.Dialog | undefined}
 */
let authDialog = undefined;

let OfficeRT = undefined;

/**
 * function to determine if an error string is really JSON
 * @param { string } str
 */
function isJsonString(str) {
  try {
    JSON.parse(str);
    return true;
  } catch (error) {
    return false;
  }
}

/**
 * function stringify a error message only if it is JSON
 * @param { string } err
 */
function stringifyError(err) {
  if (isJsonString(err)) {
    return JSON.stringify(err);
  } else {
    return err;
  }
}

// cache for access token to avoid throttling
let ATCache = undefined;

/**
 * Check if the token is nearing expiry (adjust threshold as needed)
 * @param { string } token
 */
function isTokenNearingExpiry(token) {
  const decoded = JSON.parse(atob(token.split('.')[1]));
  const expiry = decoded.exp;
  const eagerFetchThreshold = Math.floor(Date.now() / 1000) + 30; // 30 seconds before expiry
  console.debug('isTokenNearingExpiry: expiry: ', expiry, 'threshold: ', eagerFetchThreshold);
  return eagerFetchThreshold >= expiry;
}

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
  
  if (ATCache === undefined || isTokenNearingExpiry(ATCache)) {
    ATCache = await OfficeRT.auth.getAccessToken(options);
    console.debug('getAccessToken: got new token');
  } else {
    console.debug('getAccessToken: using cached token');
  }
  return ATCache;
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
  try {
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
  } catch (err) {
    // @ts-ignore
    const errStr = stringifyError(err);
    console.log(`Error: (showStatus) ${errStr}`);
  }
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

  $('<table/>', {
    width: '100%',
  }).append(
    $('<tr/>').append(
      $('<td/>', {
        class: 'ms-fontSize-24 ms-fontWeight-semibold',
        text: 'Select song to add'
      }),
      $('<td/>', {
        style: 'text-align: right',
      }).append(
        $('<i/>', {
          class: 'fas fa-rotate',
          rowspan: 2,
        }).on('click', showSecondUi)
      )
    ),
    $('<tr/>').append(
      $('<td/>', {
        class: 'ms-fontSize-12 ms-fontWeight-semibold',
        text: ' (at ActiveCell)'
      })
    )
  ).appendTo('.container');

  // Create the basefolder form
  $('<form/>').on('change', getSongCategories)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Base Folder'
    }).append($('<input/>', {
      class: 'form-input',
      type: 'text',
      required: true,
      id: 'baseFolder'
    }))).appendTo('.container');

  // Create the Song Category input form
  $('<form/>').on('change', getSongOptions)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Song Category'
    }).append($('<select/>', {
      id: 'categoryFolder',
      class: 'form-input',
      type: 'text',
      required: true,
    }))).appendTo('.container');
  
  // Create the input form
  $('<form/>').on('submit', getSongLinks)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Song'
    }).append($('<select/>', {
      id: 'songName',
      class: 'form-input',
      type: 'text',
      required: true,
    })).append($('<input/>', {
      class: 'primary-button',
      type: 'submit',
      id: 'importButton',
      value: 'Add Song'
    }))).appendTo('.container');

  $('<hr/>').appendTo('.container');
}
// </MainUiSnippet>

// <SecondUiSnippet>
function showSecondUi() {
  $('.container').empty();

  $('<table/>', {
    width: '100%'
  }).append(
    $('<tr/>').append(
      $('<td/>', {
        class: 'ms-fontSize-24 ms-fontWeight-semibold',
        text: 'Select file to add'
      }),
      $('<td/>', {
        style: 'text-align: right',
      }).append(
        $('<i/>', {
          class: 'fas fa-rotate',
          rowspan: 2,
        }).on('click', showMainUi)
      )
    ),
    $('<tr/>').append(
      $('<td/>', {
        class: 'ms-fontSize-12 ms-fontWeight-semibold',
        text: ' (at ActiveCell)'
      })
    )
  ).appendTo('.container');

  // Create the basefolder form
  $('<form/>').on('change', getItemLink)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Filename (with full path)'
    }).append($('<input/>', {
      class: 'form-input',
      type: 'text',
      required: true,
      id: 'itemPath'
    }))).appendTo('.container');

  $('<table/>').append(
    $('<tr/>').append(
      $('<td/>', {
        class: 'ms-fontSize-24 ms-fontWeight-semibold',
        text: 'Select directory listing to add'
      }),
    ),
    $('<tr/>').append(
      $('<td/>', {
        class: 'ms-fontSize-12 ms-fontWeight-semibold',
        text: ' (at ActiveCell)'
      })
    )
  ).appendTo('.container');

  // Create the basefolder form
  $('<form/>').on('change', getFolderContents)
    .append($('<label/>', {
      class: 'ms-fontSize-16 ms-fontWeight-semibold',
      text: 'Folder (with full path)'
    }).append($('<input/>', {
      class: 'form-input',
      type: 'text',
      required: true,
      id: 'folderPath'
    }))).appendTo('.container');  

  $('<hr/>').appendTo('.container');
}
// </SecondUiSnippet>

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
    // @ts-ignore
    const errStr = stringifyError(err);
    console.log(`Error: (getSongCategories) ${errStr}`);
    showStatus(`Exception populating Song Category List from OneDrive: ${errStr}`, true);
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
    // @ts-ignore
    const errStr = stringifyError(err);
    console.log(`Error: (getSongOptions) ${errStr}`);
    showStatus(`Exception populating Song List from OneDrive: ${errStr}`, true);
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
      if (files.length > 0) await WriteUrlsToSheet(files);
      showStatus(`Inserted ${files.length} song file links`, false);
    } else {
      const error = await response.json();
      showStatus(`Error getting links from OneDrive: ${JSON.stringify(error)}`, true);
    }

  } catch (err) {
    // @ts-ignore
    const errStr = stringifyError(err);
    console.log(`Error: (getSongLinks) ${errStr}`);
    showStatus(`Exception getting links from OneDrive: ${errStr}`, true);
  }
  toggleOverlay(false);
}
// <!<getSongLinksSnippet>

// <getItemLinkSnippet>
/**
 * @param {{ preventDefault: () => void; }} evt
 */
async function getItemLink(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  try {
    const apiToken = await getAccessToken({ allowSignInPrompt: true });
 
    const itemPath = $('#itemPath').val();

    const requestUrl =
      `${getBaseUrl()}/graph/itemurl?itemPath=${itemPath}`;

    const response = await fetch(requestUrl, {
      headers: {
        authorization: `Bearer ${apiToken}`
      }
    });

    if (response.ok) {
      const files = await response.json();
      if (files.length > 0) await WriteUrlsToSheet(files);
      showStatus(`Inserted ${files.length} file link`, false);
    } else {
      const error = await response.json();
      showStatus(`Error getting link from OneDrive: ${JSON.stringify(error)}`, true);
    }

  } catch (err) {
    // @ts-ignore
    const errStr = stringifyError(err);
    console.log(`Error: (getItemLink) ${errStr}`);
    showStatus(`Exception getting link from OneDrive: ${errStr}`, true);
  }
  toggleOverlay(false);
}
// </getItemLinkSnippet>

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
    const errStr = stringifyError(err);
    console.log(`Error: (WriteUrlsToSheet) ${errStr}`);
    showStatus(`Exception writing Urls to Sheet: ${errStr}`, true);
  });
}
// </WriteUrlsToSheetSnippet>

// <getFolderContentsSnippet>
/**
 *  @param {{ preventDefault: () => void; }} evt
 */
async function getFolderContents(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  console.debug('getFolderContents ...');

  try {
    const apiToken = await getAccessToken({ allowSignInPrompt: true });
    const baseFolder = $('#folderPath').val();
    const requestUrl =
      `${getBaseUrl()}/graph/folderchildren?baseFolder=${baseFolder}`;

    const response = await fetch(requestUrl, {
      headers: {
        authorization: `Bearer ${apiToken}`
      }
    });

    if (response.ok) {
      const files = await response.json();
      if (files.length > 0) await WriteFolderContentsToSheet(files);
      showStatus(`Inserted ${files.length} entries`, false);
    } else {
      const error = await response.json();
      showStatus(`Error getting listing from OneDrive: ${JSON.stringify(error)}`, true);
    }
    
  } catch (err) {
    // @ts-ignore
    const errStr = stringifyError(err);
    console.log(`Error: (getFolderContents) ${errStr}`);
    showStatus(`Exception populating folder contents from OneDrive: ${errStr}`, true);
  }

  toggleOverlay(false);
}
// </getFolderContentsSnippet>

// <WriteFolderContentsToSheetSnippet>
/**
 * @param {any[]} items
 */
async function WriteFolderContentsToSheet(items) 
{
  console.debug(`in WriteFolderContentsToSheet: items count = ${items.length}`);
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
    usedRange.load('rowCount');

    await context.sync();

    //console.debug('Used Range: ', usedRange);

    if (!usedRange.isNullObject) {
      // the sheet is not blank so make sure the row we are inserting
      // into is empty starting at the ActiveCell 
      // get the range to clear 
      let rangeToClear = sheet.getRangeByIndexes(cell.rowIndex, cell.columnIndex, items.length + 2, 3);

      //console.debug('Range to clear: ', rangeToClear);
      rangeToClear.clear();
      await context.sync();
    }

    //get the full address of the active cell
    var activeCellAddress = cell.address;

    // calculate the range needed to insert the folder listing
    var range = sheet.getRange(activeCellAddress).getResizedRange(items.length, 2);
    range.load('rowCount');
    range.load('columnCount');
    range.load('cellCount');
    await context.sync();

    //console.debug('Range to insert: ', range);
    let c1 = range.getCell(0,0);
    // @ts-ignore
    c1.values = 'Name';
    c1.format.font.bold = true;
    c1.format.horizontalAlignment = 'Center';

    let c2 = range.getCell(0, 1);
    // @ts-ignore
    c2.values = 'Type';
    c2.format.font.bold = true;
    c2.format.horizontalAlignment = 'Center';

    let c3 = range.getCell(0, 2);
    // @ts-ignore
    c3.values = 'Children';
    c3.format.font.bold = true;
    c3.format.horizontalAlignment = 'Center';

    for (var i = 0; i < items.length; i++) {
      let cell1 = range.getCell(i+1, 0);
      let cell2 = range.getCell(i+1, 1);
      let cell3 = range.getCell(i+1, 2);
      cell1.values = items[i].Name;
      cell2.values = items[i].Type;
      cell3.values = items[i].ChildCount;
    }
    range.format.autofitColumns();

  }).catch (function(err) {
    const errStr = stringifyError(err);
    console.log(`Error: (WriteFolderContentsToSheet) ${errStr}`);
    showStatus(`Exception writing Folder contents in Excel: ${errStr}`, true);
  });
}
// </WriteFolderContentsToSheetSnippet>


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

      try {
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
      } catch (error) {
        // @ts-ignore
        const errStr = stringifyError(error);
        console.log(`Error: (Office.onReady) ${errStr}`);
        showStatus(`Exception checking auth/status ${errStr}`, true);
      }
    });
  }
});
// </OfficeReadySnippet>