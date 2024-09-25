import { defineConfig } from "vite";
import { viteStaticCopy } from "vite-plugin-static-copy";
import { createHtmlPlugin } from "vite-plugin-html";
import react from "@vitejs/plugin-react";
import path from "path";

export default defineConfig(async () => {
  const { getHttpsServerOptions: getLocalHttpsOptions } = await import(
    "office-addin-dev-certs"
  );
  const httpsOptions = await getLocalHttpsOptions();

  return {
    plugins: [
      createHtmlPlugin({
        minify: false,
        pages: [
          {
            filename: "index.html",
            template: "src/index.html",
            injectOptions: {
              data: {
                injectScripts: ["/polyfill.js", "/vendor.js", "/index.js"],
              },
            },
          },
          {
            filename: "taskpane.html",
            template: "src/taskpane/taskpane.html",
            injectOptions: {
              data: {
                injectScripts: ["/polyfill.js", "/vendor.js", "/taskpane.js"],
              },
            },
          },
        ],
      }),

      react(),

      viteStaticCopy({
        targets: [
          { src: "src/taskpane/assets/*", dest: "assets" },
          {
            src: "manifest*.xml",
            dest: "",
            transform: (content) => content,
          },
        ],
      }),
    ],

    build: {
      rollupOptions: {
        input: {
          index: path.resolve(__dirname, "src/index.html"),
          taskpane: path.resolve(__dirname, "src/taskpane/index.jsx"),
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

    test: {
      environment: "jsdom",
      setupFiles: "./src/setupTests.js",
      globals: true,
      coverage: {
        exclude: [
          "postcss.config.js",
          "tailwind.config.js",
          "**/index.jsx",
          ".eslintrc.cjs",
        ],
      },
    },
  };
});
