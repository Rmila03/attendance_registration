const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Configurar carpeta pública
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Ruta para procesar registro
app.post('/registrar', (req, res) => {
    const { dni, nombre, telefono, tipo, observaciones } = req.body;
    const fechaEntrada = new Date().toISOString();

    const nuevoRegistro = {
        DNI: dni,
        Nombre: nombre,
        Telefono: telefono,
        Tipo: tipo,
        Observaciones: observaciones || '',
        FechaEntrada: fechaEntrada,
        FechaSalida: '',
        Estado: 'Activo'
    };

    const filePath = path.join(__dirname, 'asistencias.xlsx');
    let workbook;
    let trabajadores = [];
    let visitantes = [];

    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
        if (workbook.SheetNames.includes('Trabajadores')) {
            const worksheet = workbook.Sheets['Trabajadores'];
            trabajadores = xlsx.utils.sheet_to_json(worksheet);
        }
        if (workbook.SheetNames.includes('Visitantes')) {
            const worksheet = workbook.Sheets['Visitantes'];
            visitantes = xlsx.utils.sheet_to_json(worksheet);
        }
    } else {
        workbook = xlsx.utils.book_new();
    }

    // Agregar el registro a la hoja correspondiente
    if (tipo === 'Trabajador') {
        trabajadores.push(nuevoRegistro);
    } else {
        visitantes.push(nuevoRegistro);
    }

    // Crear o actualizar las hojas
    const trabajadoresWS = xlsx.utils.json_to_sheet(trabajadores);
    const visitantesWS = xlsx.utils.json_to_sheet(visitantes);

    // Eliminar hojas existentes si las hay
    if (workbook.SheetNames.includes('Trabajadores')) {
        xlsx.utils.book_remove_sheet(workbook, 'Trabajadores');
    }
    if (workbook.SheetNames.includes('Visitantes')) {
        xlsx.utils.book_remove_sheet(workbook, 'Visitantes');
    }

    // Agregar las nuevas hojas
    xlsx.utils.book_append_sheet(workbook, trabajadoresWS, 'Trabajadores');
    xlsx.utils.book_append_sheet(workbook, visitantesWS, 'Visitantes');

    // Guardar el archivo
    xlsx.writeFile(workbook, filePath);

    res.json({ message: 'Registro exitoso' });
});

// Ruta para registrar salida
app.post('/registrar-salida', (req, res) => {
    const { dni } = req.body;
    const horaSalida = new Date().toISOString();
    const filePath = path.join(__dirname, 'asistencias.xlsx');

    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'No hay registros' });
    }

    try {
        const workbook = xlsx.readFile(filePath);
        let registroEncontrado = false;

        // Buscar en ambas hojas
        ['Trabajadores', 'Visitantes'].forEach(sheetName => {
            if (workbook.SheetNames.includes(sheetName)) {
                const worksheet = workbook.Sheets[sheetName];
                const data = xlsx.utils.sheet_to_json(worksheet);
                
                const registroIndex = data.findIndex(p => p.DNI === dni && p.Estado === 'Activo');
                
                if (registroIndex !== -1) {
                    data[registroIndex].FechaSalida = horaSalida;
                    data[registroIndex].Estado = 'Completado';
                    registroEncontrado = true;

                    // Actualizar la hoja
                    const newWorksheet = xlsx.utils.json_to_sheet(data);
                    xlsx.utils.book_remove_sheet(workbook, sheetName);
                    xlsx.utils.book_append_sheet(workbook, newWorksheet, sheetName);
                }
            }
        });

        if (!registroEncontrado) {
            return res.status(404).json({ error: 'No se encontró registro activo para este DNI' });
        }

        // Guardar el archivo actualizado
        xlsx.writeFile(workbook, filePath);

        res.json({ message: 'Salida registrada exitosamente' });
    } catch (error) {
        console.error('Error al procesar la salida:', error);
        res.status(500).json({ error: 'Error al procesar la salida. Por favor, intente nuevamente.' });
    }
});

// Ruta para buscar por DNI
app.get('/buscar/:dni', (req, res) => {
    const { dni } = req.params;
    const filePath = path.join(__dirname, 'asistencias.xlsx');

    if (!fs.existsSync(filePath)) {
        return res.json({ found: false });
    }

    const workbook = xlsx.readFile(filePath);
    let persona = null;

    // Buscar en ambas hojas
    ['Trabajadores', 'Visitantes'].forEach(sheetName => {
        if (workbook.SheetNames.includes(sheetName)) {
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet);
            const encontrado = data.find(p => p.DNI === dni && p.Estado === 'Activo');
            if (encontrado) {
                persona = encontrado;
            }
        }
    });

    if (persona) {
        res.json({ found: true, persona });
    } else {
        res.json({ found: false });
    }
});

// Ruta para obtener todos los registros
app.get('/api/registros', (req, res) => {
    const filePath = path.join(__dirname, 'asistencias.xlsx');

    if (!fs.existsSync(filePath)) {
        return res.json([]);
    }

    const workbook = xlsx.readFile(filePath);
    let todosLosRegistros = [];

    // Combinar registros de ambas hojas
    ['Trabajadores', 'Visitantes'].forEach(sheetName => {
        if (workbook.SheetNames.includes(sheetName)) {
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet);
            todosLosRegistros = todosLosRegistros.concat(data);
        }
    });

    res.json(todosLosRegistros);
});

app.listen(port, () => {
    console.log(`Servidor corriendo en http://localhost:${port}`);
});
