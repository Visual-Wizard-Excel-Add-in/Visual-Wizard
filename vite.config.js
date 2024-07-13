import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteStaticCopy } from "vite-plugin-static-copy";
import { createHtmlPlugin } from "vite-plugin-html";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import path from "path";

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/";

export default defineConfig(async ({ mode }) => {
  const dev = mode === "development";
  const httpsOptions = await getHttpsServerOptions();

  return {
    plugins: [
      createHtmlPlugin({
        minify: !dev,
        pages: [
          {
            filename: "index.html",
            template: "src/index.html",
            injectOptions: {
              data: {
                injectScripts: dev
                  ? ["/polyfill.js", "/vendor.js", "/index.js"]
                  : ["/polyfill.js", "/vendor.js", "/index.js"],
              },
            },
          },
          {
            filename: "taskpane.html",
            template: "src/taskpane/taskpane.html",
            injectOptions: {
              data: {
                injectScripts: dev
                  ? ["/polyfill.js", "/vendor.js", "/taskpane.js"]
                  : ["/polyfill.js", "/vendor.js", "/taskpane.js"],
              },
            },
          },
          {
            filename: "commands.html",
            template: "src/commands/commands.html",
            injectOptions: {
              injectScripts: ["/commands.js"],
            },
          },
        ],
      }),

      react(),

      viteStaticCopy({
        targets: [
          { src: "assets/*", dest: "assets" },
          {
            src: "manifest*.xml",
            dest: "",
            transform: (content) =>
              dev
                ? content
                : content.toString().replace(new RegExp(urlDev, "g"), urlProd),
          },
        ],
      }),

      {
        name: "debug-plugin",
        resolveId(source) {
          console.log("Resolving:", source);
          return null;
        },
      },
    ],

    build: {
      rollupOptions: {
        input: {
          index: path.resolve(__dirname, "src/index.html"),
          taskpane: path.resolve(__dirname, "src/taskpane/index.jsx"),
          commands: path.resolve(__dirname, "src/commands/commands.js"),
        },
        output: {
          entryFileNames: "[name].js",
          chunkFileNames: "[name].js",
          assetFileNames: "assets/[name][extname]",
          manualChunks: {
            vendor: ["react", "react-dom"],
          },
        },
      },
      outDir: "dist",
      sourcemap: true,
    },

    server: {
      https: httpsOptions,
      port: 3000,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };
});
