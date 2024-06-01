/* eslint-disable no-undef */

const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const nodeExternals = require('webpack-node-externals');
const path = require('path');

const { envLocal, envAzure } = require('./privateEnvOptions.js');

module.exports = async (env, options) => {
  const localhost = env.app_deploy === 'localhost';
  const dev = options.mode === 'development';
  const config = [
    {
      devtool: 'source-map',
      /*output: {
        clean: true,
      },*/
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
            {
              from: 'package.json',
              to: 'package.json',
            },
            {
              from: './manifest/manifest.xml',
              to: '[name]' + '[ext]',
              transform(content) {
                if (localhost) {
                  return content;
                } else {
                  return content.toString()
                    .replace(new RegExp(envLocal.ManifestGUID, 'g'), envAzure.ManifestGUID)
                    .replace(new RegExp(envLocal.DisplayName, 'g'), envAzure.DisplayName)
                    .replace(new RegExp(envLocal.ClientId, 'g'), envAzure.ClientId)
                    .replace(new RegExp(envLocal.Url, 'g'), envAzure.Url);
                }
              },
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
        //clean: true,
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
              from: '.env',
              to: '.',
              transform(content) {
                if (localhost) {
                  return content;
                } else {
                  return content.toString()
                    .replace(new RegExp(envLocal.ClientId, 'g'), envAzure.ClientId)
                    .replace(new RegExp(envLocal.SecretValue, 'g'), envAzure.SecretValue)
                    .replace(new RegExp(envLocal.AppDeploy, 'g'), envAzure.AppDeploy)
                    .replace(new RegExp(envLocal.Port, 'g'), envAzure.Port);
                }
              }
            },
          ],
        }),
      ],
    },
  ];
  
  return config;
};
  
