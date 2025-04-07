/* eslint-disable no-undef */
const path = require("path");
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
          {
            from: "./user-manual.html",
            to: "[name][ext]",
          },
          {
            from: "manifests/manifest.*.xml",
            to: "manifests/[name][ext]",
            transform(content, filename) {
              console.log(`Transforming ${filename}. Dev mode:`, dev);
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
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
    },
  };

  return config;
};
