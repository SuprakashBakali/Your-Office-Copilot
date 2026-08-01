/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
const webpack = require("webpack");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

// ---- Dev proxy hardening (mirrors api/proxy.js) ----
const ALLOWED_HOSTS = new Set([
  "api.openai.com",
  "api.anthropic.com",
  "generativelanguage.googleapis.com",
  "api.groq.com",
  "openrouter.ai",
  "integrate.api.nvidia.com",
  "localhost",
  "127.0.0.1",
]);

function isAllowedTarget(targetUrl) {
  let url;
  try {
    url = new URL(targetUrl);
  } catch {
    return false;
  }
  if (url.protocol === "https:") return ALLOWED_HOSTS.has(url.hostname);
  if (url.protocol === "http:" && (url.hostname === "localhost" || url.hostname === "127.0.0.1")) return true;
  return false;
}

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
    optimization: {
      splitChunks: {
        chunks: "all",
        // Split large vendor libraries into their own cacheable chunks so
        // that an app-code change doesn't invalidate the user's cache for
        // React / FluentUI (which change rarely).
        cacheGroups: {
          reactVendor: {
            test: /[\\/]node_modules[\\/](react|react-dom|scheduler)[\\/]/,
            name: "vendor-react",
            chunks: "all",
          },
          fluentuiVendor: {
            test: /[\\/]node_modules[\\/]@fluentui[\\/]/,
            name: "vendor-fluentui",
            chunks: "all",
          },
          markdownVendor: {
            test: /[\\/]node_modules[\\/](react-markdown|remark-gfm|unified|micromark|remark-parse|mdast-util-.*|micromark-.*|hast-util-.*|unist-util-.*|prism-react-renderer|prismjs)[\\/]/,
            name: "vendor-markdown",
            chunks: "all",
          },
        },
      },
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".json", ".css"],
      fallback: {
        zlib: false,
        fs: false,
        path: false,
        crypto: false,
        stream: false,
        http: false,
        https: false,
        os: false,
        net: false,
        tls: false,
        child_process: false,
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: [
            {
              loader: "ts-loader",
              options: { transpileOnly: true },
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
          generator: { filename: "assets/[name][ext]" },
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
          { from: "assets", to: "assets", noErrorOnMissing: true },
          { from: "manifests", to: "manifests", noErrorOnMissing: true },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      setupMiddlewares: (middlewares, devServer) => {
        if (!devServer) throw new Error("webpack-dev-server is not defined");

        const express = require("express");
        devServer.app.post("/api/proxy", express.json(), async (req, res) => {
          try {
            const { targetUrl, headers, body } = req.body || {};
            if (!targetUrl || typeof targetUrl !== "string") {
              return res.status(400).json({ error: "targetUrl required" });
            }
            if (!isAllowedTarget(targetUrl)) {
              return res.status(403).json({
                error: "Target URL not allowed. The proxy only permits requests to known AI provider hosts over HTTPS.",
              });
            }

            // Sanitize forwarded headers — never forward Host, Content-Length, etc.
            const safeHeaders = {};
            if (headers && typeof headers === "object") {
              for (const [k, v] of Object.entries(headers)) {
                const lk = k.toLowerCase();
                if (lk === "host" || lk === "content-length" || lk === "connection") continue;
                safeHeaders[k] = v;
              }
            }

            const fetchOptions = {
              method: "POST",
              headers: safeHeaders,
              body: typeof body === "string" ? body : JSON.stringify(body),
            };

            const response = await fetch(targetUrl, fetchOptions);
            res.status(response.status);

            response.headers.forEach((value, key) => {
              const lowerKey = key.toLowerCase();
              if (
                lowerKey !== "transfer-encoding" &&
                lowerKey !== "content-encoding" &&
                lowerKey !== "content-length" &&
                lowerKey !== "connection"
              ) {
                res.setHeader(key, value);
              }
            });

            if (!response.body) return res.end();

            const reader = response.body.getReader();
            const decoder = new TextDecoder("utf-8");
            let bytesSeen = 0;

            while (true) {
              const { done, value } = await reader.read();
              if (done) break;
              bytesSeen += value.byteLength;
              if (bytesSeen > 8 * 1024 * 1024) {
                await reader.cancel();
                res.end();
                return;
              }
              res.write(decoder.decode(value));
            }
            res.end();
          } catch (e) {
            console.error("Proxy dev error:", e);
            res.status(502).json({ error: e.message });
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
      static: { directory: path.join(__dirname, "public") },
      allowedHosts: "all",
    },
  };

  return config;
};
