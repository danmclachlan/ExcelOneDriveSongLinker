// Copyright (c) Daniel R. McLachlan.
// Licensed under the MIT License.
/* eslint-disable no-undef */

// eslint-disable-next-line @typescript-eslint/no-var-requires
const CopyWebpackPlugin = require('copy-webpack-plugin');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const HtmlWebpackPlugin = require('html-webpack-plugin');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const nodeExternals = require('webpack-node-externals');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const path = require('path');

// eslint-disable-next-line @typescript-eslint/no-var-requires
const { envDefault, envLocal, envAzure } = require('./privateEnvOptions.js');

// Function to set the output directory based on the type of deployment (localhost or azure)
// Defaults to localhost if the app_deploy environment variable is not set.
function deployLocation(env, extension) {
  var base = 'dist';
  if (env.app_deploy !== undefined) {
    base = path.join(base, env.app_deploy);
  } else {
    base = path.join(base, 'localhost');
  }
  if (extension !== null) {
    base = path.join(base, extension);
  }
  return base;
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
module.exports = async (env, _options) => {
  const localhost = env.app_deploy === 'localhost';
  const config = [
    {
      devtool: 'source-map',
      output: {
        path: path.resolve(__dirname, deployLocation(env, 'addin')),
        clean: true,
      },
      entry: {
        taskpane: ['./src/addin/taskpane.js', './src/addin/taskpane.html' ],
        consent: ['./src/addin/consent.js', './src/addin/consent.html', './src/addin/config.js'],
      },
      module: {
        rules: [
          {
            test: /\.[jt]s$/,
            exclude: /node_modules/,
            use:  ['esbuild-loader', path.resolve(__dirname,`environLoader.js?app_deploy=${env.app_deploy}`) ],
          },
          {
            test: /\.html$/,
            exclude: /node_modules/,
            use: 'html-loader', 
          },
          {
            test: /\.(png|jpg|jpeg|gif|ico)$/,
            type: 'asset/resource',
            generator: {
              filename: 'assets/[name][ext][query]',
            },
          }
        ],
      }, 
      plugins: [
        new HtmlWebpackPlugin({
          filename: 'taskpane.html',
          template: './src/addin/taskpane.html',
          chunks: ['taskpane'],
        }),
        new HtmlWebpackPlugin({
          filename: 'consent.html',
          template: './src/addin/consent.html',
          chunks: ['consent'],
        }),
        
        new CopyWebpackPlugin({
          patterns: [
            {
              from: './src/addin/assets/*',
              to: 'assets/[name][ext][query]',
            },
          ],
        }),
      ],
    },
    {
      devtool: 'source-map',
      target: 'node',
      entry: {
        server: './src/server.ts',
      },
      output: {
        path: path.resolve(__dirname, deployLocation(env, null)),
      },
      externals: [nodeExternals()],
      resolve: {
        extensions: ['.ts'],
      },
      module: {
        rules: [
          {
            test: /\.[jt]s$/,
            exclude: /node_modules/,
            use: {
              loader: 'esbuild-loader',
            },
          },
        ],
      },
      plugins: [
        new CopyWebpackPlugin({
          patterns: [
            {
              from: 'template.env',
              to: '[ext]',
              transform(content) {
                if (localhost) {
                  return content.toString()
                    .replace(new RegExp(envDefault.ClientId, 'g'), envLocal.ClientId)
                    .replace(new RegExp(envDefault.SecretValue, 'g'), envLocal.SecretValue)
                    .replace(new RegExp(envDefault.TlsCertPath, 'g'), envLocal.TlsCertPath)
                    .replace(new RegExp(envDefault.TlsKeyPath, 'g'), envLocal.TlsKeyPath)
                    .replace(new RegExp(envDefault.NodeEnv, 'g'), envLocal.NodeEnv)
                    .replace(new RegExp(envDefault.AppDeploy, 'g'), envLocal.AppDeploy)
                    .replace(new RegExp(envDefault.Port, 'g'), envLocal.Port);
                } else {
                  return content.toString()
                    .replace(new RegExp(envDefault.ClientId, 'g'), envAzure.ClientId)
                    .replace(new RegExp(envDefault.SecretValue, 'g'), envAzure.SecretValue)
                    .replace(new RegExp(envDefault.NodeEnv, 'g'), envAzure.NodeEnv)
                    .replace(new RegExp(envDefault.AppDeploy, 'g'), envAzure.AppDeploy)
                    .replace(new RegExp(envDefault.Port, 'g'), envAzure.Port);
                }
              }
            },
            {
              from: 'package.json',
              to: 'package.json',
            },
            {
              from: './manifest/manifest.template.xml',
              to: 'manifest.xml',
              transform(content) {
                if (localhost) {
                  return content.toString()
                    .replace(new RegExp(envDefault.ManifestGUID, 'g'), envLocal.ManifestGUID)
                    .replace(new RegExp(envDefault.DisplayName, 'g'), envLocal.DisplayName)
                    .replace(new RegExp(envDefault.ClientId, 'g'), envLocal.ClientId)
                    .replace(new RegExp(envDefault.SupportUrl, 'g'), envLocal.SupportUrl)
                    .replace(new RegExp(envDefault.AppDeploy, 'g'), envLocal.AppDeploy)
                    .replace(new RegExp(envDefault.Url, 'g'), envLocal.Url);
                } else {
                  return content.toString()
                    .replace(new RegExp(envDefault.ManifestGUID, 'g'), envAzure.ManifestGUID)
                    .replace(new RegExp(envDefault.DisplayName, 'g'), envAzure.DisplayName)
                    .replace(new RegExp(envDefault.ClientId, 'g'), envAzure.ClientId)
                    .replace(new RegExp(envDefault.SupportUrl, 'g'), envAzure.SupportUrl)
                    .replace(new RegExp(envDefault.AppDeploy, 'g'), envAzure.AppDeploy)
                    .replace(new RegExp(envDefault.Url, 'g'), envAzure.Url);
                }
              },
            },
            {
              from: './manifest/manifest.template.xml',
              to: `../ExcelOneDriveSongLinker.${env.app_deploy}.xml`,
              transform(content) {
                if (localhost) {
                  return content.toString()
                    .replace(new RegExp(envDefault.ManifestGUID, 'g'), envLocal.ManifestGUID)
                    .replace(new RegExp(envDefault.DisplayName, 'g'), envLocal.DisplayName)
                    .replace(new RegExp(envDefault.ClientId, 'g'), envLocal.ClientId)
                    .replace(new RegExp(envDefault.SupportUrl, 'g'), envLocal.SupportUrl)
                    .replace(new RegExp(envDefault.AppDeploy, 'g'), envLocal.AppDeploy)
                    .replace(new RegExp(envDefault.Url, 'g'), envLocal.Url);
                } else {
                  return content.toString()
                    .replace(new RegExp(envDefault.ManifestGUID, 'g'), envAzure.ManifestGUID)
                    .replace(new RegExp(envDefault.DisplayName, 'g'), envAzure.DisplayName)
                    .replace(new RegExp(envDefault.ClientId, 'g'), envAzure.ClientId)
                    .replace(new RegExp(envDefault.SupportUrl, 'g'), envAzure.SupportUrl)
                    .replace(new RegExp(envDefault.AppDeploy, 'g'), envAzure.AppDeploy)
                    .replace(new RegExp(envDefault.Url, 'g'), envAzure.Url);
                }
              },
            },
          ],
        }),
      ],
    },
    {
      devtool: 'source-map',
      entry: {
        UserHelp: './src/userdocs/UserHelp.js',
        InstallationGuide: './src/userdocs/InstallationGuide.js',
      },
      output: {
        filename: '[name].js', // Output as a JavaScript file
        path: path.resolve(__dirname, deployLocation(env, 'docs')),
        clean: true,
      },
      module: {
        rules: [
          { test: /\.html$/, loader: 'html-loader'},
          {
            test: /\.md$/,
            use: [
              { 
                loader: 'markdownit-loader',
                options: {
                  preset: 'default',
                  breaks: false,
                  preprocess: function(_markdownIt, source) {
                    if (localhost) {
                      return source.toString()
                        .replace(new RegExp(envDefault.Url, 'g'), envLocal.Url)
                        .replace(new RegExp(envDefault.GitHubUrl, 'g'), envLocal.GitHubUrl);
                    } else {
                      return source.toString()
                        .replace(new RegExp(envDefault.Url, 'g'), envAzure.Url)
                        .replace(new RegExp(envDefault.GitHubUrl, 'g'), envAzure.GitHubUrl);
                    }
                  }
                },
              },
            ],
            type: 'asset/source',
          },
        ],
      },
      plugins: [
        new HtmlWebpackPlugin({
          filename: 'UserHelp.html',   // Output filename
          template: './src/userdocs/template.html',
          chunks: ['UserHelp'],
        }),
        new HtmlWebpackPlugin({
          filename: 'InstallationGuide.html',   // Output filename
          template: './src/userdocs/template.html',
          chunks: ['InstallationGuide'], 
        }),
      ]
    },
  ];
  
  return config;
};
  
