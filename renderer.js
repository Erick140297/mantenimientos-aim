const { ipcRenderer } = require("electron");
let filePath = "";

document.getElementById("load-file").addEventListener("click", async () => {
  document.getElementById("file-path").innerText = "";
  document.getElementById("process-result").innerText = "";
  const filePaths = await ipcRenderer.invoke("open-file-dialog");
  filePath = filePaths[0] || "";
  document.getElementById("file-path").innerText =
    filePath || "No se pudo seleccionar el archivo";
});

document.getElementById("process-file").addEventListener("click", async () => {
  if (!filePath) {
    document.getElementById("process-result").innerText =
      "Por favor selecciona un archivo válido primero";
    return;
  }
  const result = await ipcRenderer.invoke("process-excel-file", filePath);
  document.getElementById("process-result").innerText = result;
});
