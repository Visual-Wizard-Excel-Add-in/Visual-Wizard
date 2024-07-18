import "core-js/stable";
import "regenerator-runtime/runtime";
import { FluentProvider } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";

import App from "./components/App";
import { lightTheme, darkTheme } from "./utils/style";
import "./index.css";

const rootElement = document.getElementById("root");
const root = rootElement ? createRoot(rootElement) : undefined;

const prefersDarkScheme = window.matchMedia("(prefers-color-scheme: dark)");
const theme = prefersDarkScheme.matches ? darkTheme : lightTheme;

const globalStyles = `
  .fui-AccordionHeader__button {
    min-height: 25px !important;
    height: 25px !important;
    line-height: 25px !important;
  }
  .fui-Listbox {
    min-width: 0% !important;
    width: 6.5rem !important;
  }
  .fui-Dropdown {
    min-width: 0% !important;
    width: 6rem !important;
  }
  #fui-r1 {
    min-width: 0% !important;
    width: 6rem !important;
  }
`;

const styleSheet = document.createElement("style");
styleSheet.type = "text/css";
styleSheet.innerText = globalStyles;
document.head.appendChild(styleSheet);

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    root?.render(
      <FluentProvider theme={theme}>
        <App />
      </FluentProvider>,
    );
  }
});
