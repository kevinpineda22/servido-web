const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());

const filePath = process.env.EXCEL_FILE_PATH || 'libro.xlsx';



// Función para agregar una fila con estilos preservados
async function addRowWithStyles(data, res) {
    let workbook = new ExcelJS.Workbook();

    // Leer el archivo Excel existente o crear uno nuevo
    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
    } else {
        const worksheet = workbook.addWorksheet('Hoja1');
        worksheet.addRow(['Actividad', 'Encargado', 'Estado', 'Fecha de inicio', 'Fecha final']);
        await workbook.xlsx.writeFile(filePath);
        await workbook.xlsx.readFile(filePath);
    }

    const worksheet = workbook.getWorksheet(1);

    // Añadir nueva fila a la hoja
    const newRow = worksheet.addRow([data.activity, data.employee, data.estado, data.startDate, data.endDate]);

    // Copiar estilos de la fila anterior si existe
    if (worksheet.rowCount > 1) {
        const lastRow = worksheet.getRow(worksheet.rowCount - 1);
        lastRow.eachCell((cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);
            newCell.style = cell.style;
        });
    }

    // Escribir el archivo Excel
    try {
        await workbook.xlsx.writeFile(filePath);
        console.log('Archivo Excel actualizado');
        res.send('Actividad guardada exitosamente en el archivo Excel');
    } catch (error) {
        if (error.code === 'EBUSY') {
            console.error('Error: El archivo está siendo utilizado por otro proceso o está bloqueado.');
            res.status(500).send('Error: El archivo está siendo utilizado por otro proceso o está bloqueado.');
        } else {
            console.error('Error al escribir el archivo Excel:', error);
            res.status(500).send('Error al escribir el archivo Excel');
        }
    }
}

// Servir archivos estáticos desde la carpeta 'public'
app.use(express.static(path.join(__dirname, 'public')));

// Ruta para la página principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Ruta para manejar POST en /save-activity
app.post('/save-activity', (req, res) => {
    const data = req.body;
    console.log('Datos recibidos:', data);

    // Añadir nueva fila con la actividad y mantener estilos
    addRowWithStyles(data, res);
});

// Ruta para manejar GET en /
app.get('/', (req, res) => {
    res.send('Servidor corriendo');
});

app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
