/* eslint-disable no-undef */
const path = require("path");
const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const devServerPort = process.env.DEV_SERVER_PORT ? parseInt(process.env.DEV_SERVER_PORT, 10) : 3000;

const urlDev = `https://localhost:${devServerPort}/`;
const urlProd = "https://michel-bodje.github.io/automate/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const baseUrl = dev ? urlDev : urlProd;
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./app/taskpane.js", "./app/taskpane.html"],
      commands: "./app/commands/commands.js",
    },
    output: {
      path: path.resolve(__dirname, "docs"),
      publicPath: baseUrl,
      filename: dev ? "[name].bundle.js" : "[name].[contenthash].js" 
    },
    resolve: {
      extensions: [".html", ".js"],
      fallback: {
        process: require.resolve("process"),
      },
    },
    module: {
      rules: [
        {
          test: /\.css$/,
          use: [
            'style-loader',
            'css-loader'
          ]
        },
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/images/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./app/taskpane.html",
        chunks: ["polyfill", "taskpane"],
        inject: true,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/images/*",
            to: "assets/images/[name][ext][query]",
          },
          {
            from: "assets/templates/en/*",
            to: "assets/templates/en/[name][ext][query]",
          },
          {
            from: "assets/templates/fr/*",
            to: "assets/templates/fr/[name][ext][query]",
          },
          /*
          {
          from: "assets/templates/en/Equipe\ Allen\ Madelin\ (admin@amlex.ca)_files/*",
          to: 'assets/templates/en/Equipe\ Allen\ Madelin\ (admin@amlex.ca)_files/[name][ext][query]',
          },
          */
          {
            from: "./user-manual.html",
            to: "[name][ext]",
          },
          {
            from: "manifests/manifest.*.xml",
            to: "manifests/[name][ext]",
            transform(content, filename) {
              console.log(`Transforming ${filename}. Dev mode:`, dev);
              if (!dev) {
              // In production, do not modify the manifest (assume it's production ready)
              return content;
              } else {
              // In development, replace production URLs with localhost and update <Id> tag for dev manifest
              const isOutlook = filename.includes("manifest.outlook.xml");
              const isWord = filename.includes("manifest.word.xml");

              // Define GUIDs for each manifest and environment
              const IDS = {
                outlook: {
                  prod: "bf0f99cf-6ef7-4fa5-9ffe-7f3197a4b10a",
                  dev: "4d546269-e983-4e12-aa50-ed569039df51",
                },
                word: {
                  prod: "607de257-2aa6-4c5c-bdd4-e122d47e9d9e",
                  dev: "cc48ccc7-2d17-4a3a-9fe9-41eb9bacf7bf",
                },
              };

              const manifestType = isOutlook ? "outlook" : isWord ? "word" : null;
              const prodId = manifestType ? IDS[manifestType].prod : undefined;
              const devId = manifestType ? IDS[manifestType].dev : undefined;

              // Replace production URLs with localhost
              let result = content.toString()
                .replace(new RegExp(urlProd.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), "g"), urlDev);

              // Replace production GUID with dev GUID
              if (prodId && devId) {
                result = result.replace(new RegExp(`<Id>${prodId}<\/Id>`), `<Id>${devId}</Id>`);
              }

              return result;
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./app/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.ProvidePlugin({
        process: "process/browser",
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
        "Access-Control-Allow-Headers": "X-Requested-With, content-type, Authorization"
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: devServerPort,
      devMiddleware: {
        writeToDisk: true
      },
    },
  };

  return config;
};
