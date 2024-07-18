const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();

const app = express();
const PORT = 3816;
app.use(cors());
app.use(bodyParser.json());

const filePath = path.join('C:', 'Users', 'USER', 'Downloads', 'libro.xlsx');

// Configurar base de datos SQLite
const db = new sqlite3.Database(':memory:');
db.serialize(() => {
    db.run("CREATE TABLE activities (activity TEXT, employee TEXT, startDate TEXT, endDate TEXT, estado TEXT)");
});

function insertActivity(activity, employee, startDate, endDate, estado) {
    db.run("INSERT INTO activities (activity, employee, startDate, endDate, estado) VALUES (?, ?, ?, ?, ?)", 
        [activity, employee, startDate, endDate, estado]);
}

function getActivities(callback) {
    db.all("SELECT * FROM activities", (err, rows) => {
        callback(rows);
    });
}

// Función para agregar una fila con estilos preservados
async function addRowWithStyles(data, res) {
    let workbook = new ExcelJS.Workbook();

    // Leer el archivo Excel existente o crear uno nuevo
    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
    } else {
        const worksheet = workbook.addWorksheet('Sheet1');
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

// Función para exportar datos de la base de datos a Excel
async function exportToExcel() {
    let workbook = new ExcelJS.Workbook();

    // Leer el archivo Excel existente o crear uno nuevo
    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
    } else {
        const worksheet = workbook.addWorksheet('Sheet1');
        worksheet.addRow(['Actividad', 'Encargado', 'Estado', 'Fecha de inicio', 'Fecha final']);
        await workbook.xlsx.writeFile(filePath);
        await workbook.xlsx.readFile(filePath);
    }

    const worksheet = workbook.getWorksheet(1);

    // Obtener todas las actividades de la base de datos y agregarlas al archivo Excel
    getActivities((rows) => {
        rows.forEach((row) => {
            worksheet.addRow([row.activity, row.employee, row.estado, row.startDate, row.endDate]);
        });

        workbook.xlsx.writeFile(filePath).then(() => {
            console.log('Datos exportados a Excel');
        }).catch((error) => {
            console.error('Error al escribir el archivo Excel:', error);
        });
    });
}

// Exportar datos a Excel cada minuto
setInterval(exportToExcel, 60000);

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

    // Insertar nueva actividad en la base de datos temporalmente
    insertActivity(data.activity, data.employee, data.startDate, data.endDate, data.estado);
    res.send('Actividad registrada temporalmente');
});

// Ruta para manejar GET en /
app.get('/', (req, res) => {
    res.send('Servidor corriendo');
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Servidor corriendo en http://11.11.13.207:${PORT}`);
});
