import "./taskpane.css";
import { initializeOffice } from "@/utils/office";
import { renderApp } from "@/store/app";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeOffice();
    renderApp(document.getElementById("app"));
  }
});
