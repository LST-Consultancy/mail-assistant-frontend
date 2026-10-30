/* global document, Office */
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";
import { isDarkTheme } from "../services/office";
import { App } from "./App";

const root = createRoot(document.getElementById("container")!);

Office.onReady(() => {
  root.render(
    <FluentProvider theme={isDarkTheme() ? webDarkTheme : webLightTheme} style={{ minHeight: "100vh" }}>
      <App />
    </FluentProvider>,
  );
});
