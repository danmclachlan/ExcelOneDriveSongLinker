// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <ServerSnippet>
import express, { RequestHandler } from 'express';
import https from 'https';
import fs from 'fs';
import dotenv from 'dotenv';
import path from 'path';
import logger from 'morgan';

// Load .env file
const result = dotenv.config();
if (result.error) {
  console.error('dotenv failed: ', result.error);
} 
console.log('dotenv result: ', result.parsed);

import authRouter from './api/auth';
import graphRouter from './api/graph';

const app = express();
const PORT = process.env.PORT || 443;

app.set('port', PORT);
app.use(logger('dev'));

// Support JSON payloads
app.use(express.json() as RequestHandler);

/* Turn off caching when developing */
if (process.env.NODE_ENV !== 'production') {
  app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next();
  });

  app.use(express.static(process.cwd(), { etag: false }));
  app.use(express.static(path.join(process.cwd(), 'addin'), { etag: false }));
  app.use(express.static(path.join(process.cwd(), 'dist'), { etag: false }));
} else {
  // In production mode, let static files be cached.
  app.use(express.static(process.cwd()));
  app.use(express.static(path.join(process.cwd(), 'addin')));
  app.use(express.static(path.join(process.cwd(), 'dist')));
}


console.debug(`dirname = ${process.cwd()}`);

app.use('/auth', authRouter);
app.use('/graph', graphRouter);

if (process.env.APP_DEPLOY === 'localhost') {

  const serverOptions = {
    key: fs.readFileSync(process.env.TLS_KEY_PATH || ''),
    cert: fs.readFileSync(process.env.TLS_CERT_PATH || ''),
  };

  https.createServer(serverOptions, app).listen(PORT, () => {
    console.log(`⚡️[server]: Server is running at https://localhost:${PORT}`);
  });
} else {
  // production mode
  app.listen(PORT, () => { 
    console.log(`⚡️[server]: Server is running on ${PORT}`);
  });
}
// </ServerSnippet>
