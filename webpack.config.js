const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const devCerts = require("office-addin-dev-certs");

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  const config = {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: [],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext]",
            noErrorOnMissing: true,
          },
        ],
      }),
    ],
  };

  // Dev server with HTTPS (required by Office Add-ins)
  if (dev) {
    try {
      const httpsOptions = await devCerts.getHttpsServerOptions();
      config.devServer = {
        static: {
          directory: path.join(__dirname, "dist"),
        },
        headers: {
          "Access-Control-Allow-Origin": "*",
        },
        server: {
          type: "https",
          options: httpsOptions,
        },
        port: "auto",
        hot: true,
        onListening: function (devServer) {
          const actualPort = devServer.server.address().port;
          if (actualPort !== 3000) {
            console.warn(
              `\n⚠️  Dev server running on port ${actualPort} instead of 3000.` +
              `\n   Update manifest.xml to use https://localhost:${actualPort} for Excel sideloading.\n`
            );
          }
        },
      };
    } catch (e) {
      console.warn("Could not get dev certs, falling back to basic HTTPS:", e.message);
      config.devServer = {
        static: {
          directory: path.join(__dirname, "dist"),
        },
        headers: {
          "Access-Control-Allow-Origin": "*",
        },
        server: "https",
        port: "auto",
        hot: true,
      };
    }
  }

  return config;
};
