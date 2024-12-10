/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "";
const urlProd = "https://www.fortevento.lt"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "production";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"), // Ensure output goes to the dist folder
      filename: "[name].js", // Maintain the naming convention for JS files
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
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
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      // Generate taskpane.html
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),

      // Generate commands.html
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),

      // Generate index.html in dist
      new HtmlWebpackPlugin({
        filename: "index.html",  // This will create index.html in the dist folder
        template: "./src/taskpane/taskpane.html",  // You can use the same template or modify for index.html
        chunks: ["polyfill", "taskpane"],  // Include the necessary chunks
      }),

      // Copy templates and assets to dist
      new CopyWebpackPlugin({
        patterns: [
          {
            from: path.resolve(__dirname, "src/templates"),
            to: "templates", // Copy templates to dist/templates
          },
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      static: {
        directory: path.join(__dirname, "dist"), // Serve from dist folder
        watch: true,
      },
    },
  };

  return config;
};
