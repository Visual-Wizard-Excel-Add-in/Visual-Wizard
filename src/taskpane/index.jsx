import "core-js/stable";
import "regenerator-runtime/runtime";
import React from "react";
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
} from "@fluentui/react-components";
import { createRoot } from "react-dom/client";
import App from "./components/App";

const rootElement = document.getElementById("root");
const root = rootElement ? createRoot(rootElement) : undefined;

const prefersDarkScheme = window.matchMedia("(prefers-color-scheme: dark)");
const theme = prefersDarkScheme.matches ? webDarkTheme : webLightTheme;

Office.onReady((info) => {
  console.log("Office.js is ready", info);
  if (info.host === Office.HostType.Excel) {
    root?.render(
      <FluentProvider theme={theme}>
        <App />
      </FluentProvider>,
    );
  }
});
