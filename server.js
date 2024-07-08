const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(bodyParser.json());

const filePath = path.join('C:', 'xampp', 'htdocs', 'activities.xlsx');

// Función para intentar escribir en el archivo con manejo de errores
function tryWriteFile(workbook, filePath, res) {
    try {
        XLSX.writeFile(workbook, filePath);
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

// Ruta para manejar POST en /save-activity
app.post('/save-activity', (req, res) => {
    const data = req.body;
    console.log('Datos recibidos:', data);

    // Lee el archivo Excel existente o crea uno nuevo
    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
    } else {
        workbook = XLSX.utils.book_new();
    }

    let worksheet;
    if (workbook.SheetNames.length) {
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
        worksheet = XLSX.utils.aoa_to_sheet([['Actividad', 'Encargado', 'Estado', 'Fecha de inicio', 'Fecha final', 'Observaciones']]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    }

    // Añadir nueva fila con la actividad
    const newRow = [data.activity, data.employee, data.estado, data.startDate, data.endDate, data.observacion];
    XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });
    console.log('Nueva fila añadida:', newRow);

    // Escribir el archivo Excel
    tryWriteFile(workbook, filePath, res);
});

// Ruta para manejar GET en /
app.get('/', (req, res) => {
    res.send('Servidor corriendo');
});

app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});