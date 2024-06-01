// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import Router from 'express-promise-router';
import * as graph from '@microsoft/microsoft-graph-client';
import { Drive, DriveItem } from 'microsoft-graph';
import 'isomorphic-fetch';
import { getTokenOnBehalfOf } from './auth';
import axios from 'axios';

// <GetClientSnippet>
async function getAuthenticatedClient(authHeader: string): Promise<graph.Client> {
  const accessToken = await getTokenOnBehalfOf(authHeader);

  return graph.Client.init({
    authProvider: (done) => {
      // Call the callback with the
      // access token
      done(null, accessToken || '');
    },
    debugLogging: true
  });
}
// </GetClientSnippet>

// <GetReadOnlyLinkSnippet>
async function GetReadOnlyLink(client: graph.Client, id: string): Promise<string> {
  const requestBody = {
    type: 'view',
    scope: 'anonymous'
  };

  const path = '/me/drive/items/';
  const linkItem = await client
    .api(path.concat(id, '/createlink'))
    .post(requestBody);

  // console.debug(`GetReadOnlyLink:${linkItem.link.webUrl}`);

  return linkItem.link.webUrl;
}
// </GetReadOnlyLinkSnippet>

// <ResolveShortcutSnippet>
async function ResolveShortcut(client: graph.Client, id: string): Promise<string> {

  const path = '/me/drive/items/';
  const uriItem = await client
    .api(path.concat(id, '?select=id,@microsoft.graph.downloadUrl'))
    .get();

  const downloadUrl = uriItem['@microsoft.graph.downloadUrl'];

  let resolvedShortcut = '';

  try {
    const response = await axios.get(downloadUrl);
    const fileContent = response.data; // content of the .url file

    // extract the URL= line from the content
    const urlLineMatch = fileContent.match(/URL=(.*)/i);

    if (urlLineMatch) {
      resolvedShortcut = urlLineMatch[1].trim();
    } else {
      console.error('URL= line not found in .url file content.');
    }
  } catch (error) {
    console.error('Error downloading .url file:', error);
  }

  return resolvedShortcut;
}
// </ResolveShortcutSnippet>

const graphRouter = Router();

// <GetFolderitemurlsSnippet>
interface ItemUrlObject {
  Name: string;
  Type: string;
  WebUrl: string;
}

graphRouter.get('/folderitemsurls',
  async function(req, res) {
    const authHeader = req.headers['authorization'];

    console.debug('in function for /folderitemsurls');

    if (authHeader) {
      try {
        const client = await getAuthenticatedClient(authHeader);

        const baseFolder = req.query['baseFolder']?.toString();
        const songName = req.query['songName']?.toString();

        console.debug(`baseFolder: ${baseFolder}, songName: ${songName}`);

        const folderPath = baseFolder?.concat('/', songName || '');
        const resultArray: ItemUrlObject[]  = [];

        const drive: Drive = await client
          .api('/me/drive')
          .get();

        if (drive === null) {
          throw new Error('Could not get OneDrive root (/me/drive)');
        }
        //console.debug(drive);
        
        const basepath = '/drives/';
        const driveItemList: DriveItem[] = [];
        
        const folderItem: DriveItem = await client
          .api(basepath.concat(drive.id || '', '/root:/', folderPath || ''))
          .get();

        if (folderItem === null) {
          throw new Error(`Could not get OneDrive folder (${folderPath})`);
        }
    
        //console.debug(folderItem);

        value = await GetReadOnlyLink(client, folderItem.id || '');
        if (value !== undefined) {
          resultArray.push({  'Name': folderItem.name || '', 'Type': 'Folder', 'WebUrl': value});  
        }

        const driveItemPage: graph.PageCollection = await client
          .api(basepath.concat(drive.id || '', '/items/', folderItem.id || '', '/children'))
          .get();
        
        // Set up a PageIterator to process the events in the result
        // and request subsequent "pages" if there are more than 25
        // on the server
        const callback: graph.PageIteratorCallback = (driveItem) => {
          driveItemList.push(driveItem);
          return true;
        };
    
        const iterator = new graph.PageIterator(client, driveItemPage, callback, undefined);
        await iterator.iterate();

        // console.debug(`FileCount: ${driveItemList.length}`);
    
        for (let i = 0; i < driveItemList.length; i++) {
          if (driveItemList[i].file !== null) {
            // eslint-disable-next-line no-var
            var value: string | undefined = undefined;

            //console.debug(`driveItemList[${i}]:`);
            //console.debug(driveItemList[i]);
    
            const id = driveItemList[i].id;
            const name = driveItemList[i].name;
            if (id !== null && id !== undefined) {
              if (name !== null && name !== undefined && name.endsWith('.url')) {
                value = await ResolveShortcut(client, id);
                if (value !== '') {
                  resultArray.push({  'Name': driveItemList[i].name || '', 'Type': 'Shortcut', 'WebUrl': value});
                }
              } else {
                value = await GetReadOnlyLink(client, id);
                if (value !== undefined) {
                  resultArray.push({  'Name': driveItemList[i].name || '', 'Type': 'File', 'WebUrl': value});
                }
              }
            } 
          }
        }
        /**/
        console.debug('resultArray: ', resultArray);

        // Return the array of events
        res.status(200).json(resultArray);
      } catch (error) {
        console.log(error);
        res.status(500).json(error);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
  }
);
// </GetFolderitemurlsSnippet>

// <GetFolderChildrenSnippet>
interface ItemDescObject {
  Name: string;
  Type: string;
  ChildCount: number;
}

graphRouter.get('/folderchildren',
  async function(req, res) {
    const authHeader = req.headers['authorization'];

    console.debug('in function for /folderchildren');

    if (authHeader) {
      try {
        const client = await getAuthenticatedClient(authHeader);

        const folderPath = req.query['baseFolder']?.toString();
        
        console.debug(`baseFolder: ${folderPath}`);

        const resultArray: ItemDescObject[]  = [];

        const drive: Drive = await client
          .api('/me/drive')
          .get();

        if (drive === null) {
          throw new Error('Could not get OneDrive root (/me/drive)');
        }
        //console.debug(drive);
        
        const basepath = '/drives/';
        const driveItemList: DriveItem[] = [];
        
        const folderItem: DriveItem = await client
          .api(basepath.concat(drive.id || '', '/root:/', folderPath || ''))
          .get();

        if (folderItem === null) {
          throw new Error(`Could not get OneDrive folder (${folderPath})`);
        }
    
        // console.debug(folderItem);

        const driveItemPage: graph.PageCollection = await client
          .api(basepath.concat(drive.id || '', '/items/', folderItem.id || '', '/children'))
          .get();
        
        // Set up a PageIterator to process the events in the result
        // and request subsequent "pages" if there are more than 25
        // on the server
        const callback: graph.PageIteratorCallback = (driveItem) => {
          driveItemList.push(driveItem);
          return true;
        };
    
        const iterator = new graph.PageIterator(client, driveItemPage, callback, undefined);
        await iterator.iterate();

        // console.debug(`FileCount: ${driveItemList.length}`);
    
        for (let i = 0; i < driveItemList.length; i++) {

          // console.debug(`driveItemList[${i}]:`, driveItemList[i].name,'\nfolder: ', driveItemList[i].folder, '\nfile: ', driveItemList[i].file);

          if (driveItemList[i].file !== undefined) {
            // the item is a file
            const name = driveItemList[i].name;

            if (name !== null && name !== undefined) {
              if (name.endsWith('.url')) {
                resultArray.push({  'Name': name, 'Type': 'Shortcut', 'ChildCount': 0});
              } else {
                resultArray.push({  'Name': name, 'Type': 'File', 'ChildCount': 0});
              }
            }
          } else if (driveItemList[i].folder !== undefined) {
            // item is a folder
            const name = driveItemList[i].name;

            if (name !== null && name !== undefined) {
              resultArray.push({  'Name': name, 'Type': 'Folder', 'ChildCount': 0 });
            }
          }
        }

        /**/
        console.debug('resultArray: ', resultArray);

        // Return the array of events
        res.status(200).json(resultArray);
      } catch (error) {
        console.log(error);
        res.status(500).json(error);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
  }
);
// </GetFolderChildrenSnippet>

export default graphRouter;