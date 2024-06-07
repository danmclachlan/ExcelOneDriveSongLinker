// Copyright (c) Daniel R. McLachlan.
// Licensed under the MIT License.

const envDefault = {
  ManifestGUID: 'NEW_GUID_HERE',
  DisplayName:  'YOUR_DISPLAY_NAME_HERE',
  ClientId:     'YOUR_CLIENT_ID_HERE',
  SecretValue:  'YOUR_CLIENT_SECRET_HERE',
  TlsCertPath:  'PATH_TO_LOCALHOST.crt',
  TlsKeyPath:   'PATH_TO_LOCALHOST.key',
  NodeEnv:      'development',
  Url:          'localhost:3000',
  AppDeploy:    'APP_DEPLOY=\'localhost\'',
  Port:         '3000',
  SupportUrl:   'https://www.contoso.com/help',
  GitHubUrl:    'YOUR_GITHUB_LOCATION'
};

const envLocal = {
  ManifestGUID: '<your new local guid>',
  DisplayName:  'Excel OneDrive Song Linker (localhost)',
  ClientId:     '<your localhost CLIENTID>',
  SecretValue:  '<your localhost Client Secret>',
  TlsCertPath:  '<path to LOCALHOST.CRT>',
  TlsKeyPath:   '<path to LOCALHOST.KEY>',
  NodeEnv:      'development',
  Url:          'localhost:3000',
  AppDeploy:    'APP_DEPLOY=\'localhost\'',
  Port:         '3000',
  SupportUrl:   'https://localhost:3000/UserHelp.html',
  GitHubUrl:    '<your github url including https://>'
};
    
const envAzure = {
  ManifestGUID: '<your new Azure guid>',
  DisplayName:  'Excel OneDrive Song Linker (Azure)',
  ClientId:     '<your Azure CLIENTID>',
  SecretValue:  '<your Azure Client Secret>',
  TlsCertPath:  '',
  TlsKeyPath:   '',
  NodeEnv:      'development',
  Url:          '<App Service app URL>',
  AppDeploy:    'APP_DEPLOY=\'Azure\'',
  Port:         '443',
  SupportUrl:   'https://<App Service app URL>/UserHelp.html',
  GitHubUrl:    '<your github url including https://>'
};

exports.envDefault = envDefault;
exports.envLocal = envLocal;
exports.envAzure = envAzure;