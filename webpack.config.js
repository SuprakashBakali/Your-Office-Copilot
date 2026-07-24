/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

const urlDev = "https://localhost:3000/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

const webpack = require("webpack");

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  const config = {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "public"),
      filename: "[name].[contenthash].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".json", ".css"],
      fallback: {
        "zlib": false,
        "fs": false,
        "path": false,
        "crypto": false,
        "stream": false,
        "http": false,
        "https": false,
        "os": false,
        "net": false,
        "tls": false,
        "child_process": false,
      }
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: [
            {
              loader: "ts-loader",
              options: {
                transpileOnly: true,
              },
            },
          ],
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },
    plugins: [
      new webpack.NormalModuleReplacementPlugin(/^node:/, (resource) => {
        resource.request = resource.request.replace(/^node:/, "");
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
            noErrorOnMissing: true,
          },
          {
            from: "manifests",
            to: "manifests",
            noErrorOnMissing: true,
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      setupMiddlewares: (middlewares, devServer) => {
        if (!devServer) throw new Error('webpack-dev-server is not defined');
        
        const express = require('express');
        devServer.app.post('/api/proxy', express.json(), async (req, res) => {
          try {
            const { targetUrl, headers, body } = req.body;
            if (!targetUrl) return res.status(400).json({ error: 'targetUrl required' });

            const fetchOptions = {
              method: 'POST',
              headers: headers || {},
              body: typeof body === 'string' ? body : JSON.stringify(body),
            };

            const response = await fetch(targetUrl, fetchOptions);
            res.status(response.status);
            
            response.headers.forEach((value, key) => {
              const lowerKey = key.toLowerCase();
              if (lowerKey !== 'transfer-encoding' && lowerKey !== 'content-encoding') {
                res.setHeader(key, value);
              }
            });

            if (!response.body) return res.end();

            const reader = response.body.getReader();
            const decoder = new TextDecoder('utf-8');
            while (true) {
              const { done, value } = await reader.read();
              if (done) break;
              res.write(decoder.decode(value));
            }
            res.end();
          } catch (e) {
            console.error('Proxy dev error:', e);
            res.status(500).json({ error: e.message });
          }
        });
        return middlewares;
      },
      server: {
        type: "https",
        options: dev ? await getHttpsOptions() : {},
      },
      port: 3000,
      hot: true,
      static: {
        directory: path.join(__dirname, "public"),
      },
      allowedHosts: "all",
    },
  };

  return config;
};
