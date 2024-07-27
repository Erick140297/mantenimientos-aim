const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const XLSX = require("xlsx");
const moment = require("moment");
const fs = require("fs");

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 750,
    height: 750,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  mainWindow.loadFile("index.html");
  // mainWindow.webContents.openDevTools();

  // Ocultar la barra de menú
  mainWindow.setMenuBarVisibility(false);

  ipcMain.handle("open-file-dialog", async () => {
    const result = await dialog.showOpenDialog({
      properties: ["openFile"],
      filters: [{ name: "Excel Files", extensions: ["xls", "xlsx"] }],
    });
    return result.filePaths;
  });

  ipcMain.handle("process-excel-file", async (event, filePath) => {
    try {
      const workbook = XLSX.readFile(filePath);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      let data = XLSX.utils.sheet_to_json(worksheet);

      data = data.map((row) => {
        if (row.date) {
          let fecha;
          if (typeof row.date === "number") {
            const date = XLSX.SSF.parse_date_code(row.date);
            fecha = moment(new Date(date.y, date.m - 1, date.d));
          } else {
            fecha = moment(row.date, "DD-MM-YYYY", true);
          }

          if (!fecha.isValid()) {
            console.warn(`Fecha no válida: ${row.date}`);
            return row;
          }

          row.date = fecha.format("DD-MM-YYYY");

          if (row.duracion && !isNaN(row.duracion)) {
            const limitDate = fecha
              .clone()
              .add(parseInt(row.duracion, 10), "days");
            row.limit_date = limitDate.format("DD-MM-YYYY");
          } else {
            row.limit_date = fecha.format("DD-MM-YYYY");
          }
        }
        return row;
      });

      let newData = [];

      data.forEach((row) => {
        let fechaOriginal = moment(row.date, "DD-MM-YYYY");
        let incremento;

        switch (row.tipo) {
          case "mensual":
            incremento = 1;
            break;
          case "trimestral":
            incremento = 3;
            break;
          case "cuatrimestral":
            incremento = 4;
            break;
          case "bimestral":
            incremento = 2;
            break;
          case "semestral":
            incremento = 6;
            break;
          case "anual":
            incremento = 12;
            break;
          case "bianual":
            incremento = 24;
            break;
          case "semanal":
            incremento = "semanal";
            break;
          case "quincenal":
            incremento = "quincenal";
            break;
          default:
            console.warn(`Tipo desconocido: ${row.tipo}`);
            return;
        }

        if (row.tipo === "semestral") {
          for (let i = 0; i < 2; i++) {
            let fechaConMesIncrementado = fechaOriginal
              .clone()
              .add(i * incremento, "months");
            let nuevoRegistro = { ...row };
            nuevoRegistro.date = fechaConMesIncrementado.format("DD-MM-YYYY");

            if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
              const limitDate = fechaConMesIncrementado
                .clone()
                .add(parseInt(nuevoRegistro.duracion, 10), "days");
              nuevoRegistro.limit_date = limitDate.format("DD-MM-YYYY");
            }

            newData.push(nuevoRegistro);
          }
        } else if (incremento === "semanal") {
          let fechaConSemanaIncrementada = fechaOriginal.clone();
          while (fechaConSemanaIncrementada.year() === fechaOriginal.year()) {
            let nuevoRegistro = { ...row };
            nuevoRegistro.date =
              fechaConSemanaIncrementada.format("DD-MM-YYYY");

            if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
              const limitDate = fechaConSemanaIncrementada
                .clone()
                .add(parseInt(nuevoRegistro.duracion, 10), "days");
              nuevoRegistro.limit_date = limitDate.format("DD-MM-YYYY");
            }

            newData.push(nuevoRegistro);
            fechaConSemanaIncrementada.add(1, "week");
          }
        } else if (incremento === "quincenal") {
          for (let mes = fechaOriginal.month(); mes <= 11; mes++) {
            let fechaConMesIncrementado = fechaOriginal.clone().month(mes);
            if (fechaConMesIncrementado.year() === fechaOriginal.year()) {
              let nuevoRegistro = { ...row };
              nuevoRegistro.date = fechaConMesIncrementado.format("DD-MM-YYYY");

              if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
                const limitDate = fechaConMesIncrementado
                  .clone()
                  .add(parseInt(nuevoRegistro.duracion, 10), "days");
                nuevoRegistro.limit_date = limitDate.format("DD-MM-YYYY");
              }

              newData.push(nuevoRegistro);

              let fechaCon15DiasAdicionales = fechaConMesIncrementado
                .clone()
                .add(15, "days");
              if (fechaCon15DiasAdicionales.year() === fechaOriginal.year()) {
                let nuevoRegistroCon15Dias = { ...row };
                nuevoRegistroCon15Dias.date =
                  fechaCon15DiasAdicionales.format("DD-MM-YYYY");

                if (
                  nuevoRegistroCon15Dias.duracion &&
                  !isNaN(nuevoRegistroCon15Dias.duracion)
                ) {
                  const limitDate = fechaCon15DiasAdicionales
                    .clone()
                    .add(parseInt(nuevoRegistroCon15Dias.duracion, 10), "days");
                  nuevoRegistroCon15Dias.limit_date =
                    limitDate.format("DD-MM-YYYY");
                }

                newData.push(nuevoRegistroCon15Dias);
              }
            }
          }
        } else {
          let fechaFinDelAño = fechaOriginal.clone().endOf("year");
          let fechaConIncremento = fechaOriginal.clone();

          while (fechaConIncremento.isSameOrBefore(fechaFinDelAño)) {
            let nuevoRegistro = { ...row };
            nuevoRegistro.date = fechaConIncremento.format("DD-MM-YYYY");

            if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
              const limitDate = fechaConIncremento
                .clone()
                .add(parseInt(nuevoRegistro.duracion, 10), "days");
              nuevoRegistro.limit_date = limitDate.format("DD-MM-YYYY");
            }

            newData.push(nuevoRegistro);

            if (row.tipo === "bianual") {
              fechaConIncremento.add(2 * incremento, "months");
            } else {
              fechaConIncremento.add(incremento, "months");
            }
          }

          if (row.tipo === "bianual") {
            let fechaConIncrementoAdicional = fechaOriginal
              .clone()
              .add(incremento, "months");
            if (fechaConIncrementoAdicional.year() !== fechaOriginal.year()) {
              let nuevoRegistroAdicional = { ...row };
              nuevoRegistroAdicional.date =
                fechaConIncrementoAdicional.format("DD-MM-YYYY");

              if (
                nuevoRegistroAdicional.duracion &&
                !isNaN(nuevoRegistroAdicional.duracion)
              ) {
                const limitDate = fechaConIncrementoAdicional
                  .clone()
                  .add(parseInt(nuevoRegistroAdicional.duracion, 10), "days");
                nuevoRegistroAdicional.limit_date =
                  limitDate.format("DD-MM-YYYY");
              }

              newData.push(nuevoRegistroAdicional);
            }
          }
        }
      });

      newData = newData.map(({ duracion, tipo, ...resto }) => resto);

      // Guardar el archivo XLSX
      const newWorksheet = XLSX.utils.json_to_sheet(newData, {
        header: ["infrastructure_id", "date", "limit_date", "description"],
      });
      workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
      const xlsxOutputPath = filePath.replace(
        /(\.[\w\d_-]+)$/i,
        "_processed.xlsx"
      );
      XLSX.writeFile(workbook, xlsxOutputPath);

      // Convertir los datos a formato TSV y guardarlo
      const tsvData = convertToTSV(newData);
      const tsvOutputPath = filePath.replace(
        /(\.[\w\d_-]+)$/i,
        "_processed.tsv"
      );
      fs.writeFileSync(tsvOutputPath, tsvData);

      return `Archivos procesados y guardados correctamente como: ${xlsxOutputPath} y ${tsvOutputPath}`;
    } catch (error) {
      return `Error al procesar el archivo: ${error.message}`;
    }
  });
}

function convertToTSV(data) {
  const headers = Object.keys(data[0]);
  const tsvRows = data.map((row) =>
    headers.map((field) => row[field] || "").join("\t")
  );
  return [headers.join("\t"), ...tsvRows].join("\n");
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
