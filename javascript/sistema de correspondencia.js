// Sistema DEU – generación automatizada de oficios
// (basado en test2.pdf, versión 7.3)

// ==================== CONFIGURACIÓN GLOBAL ====================
const CONFIG = {
  PLANTILLA: `{{codigo}}

Caracas, {{fecha}}

Ciudadano:
{{director}}

{{cargo}} {{escuela}} Universidad Central de Venezuela Presente.

Asunto: {{asunto}}

Sin más a que referirme, queda de usted.

Atentamente.

Profa. Mercy Ospina Directora de Extensión Universitaria

“Ciudad Universitaria de Caracas, Patrimonio de la Humanidad: 2025, vigésimo quinto aniversario de la Declaración de la UNESCO ”

Edif. Biblioteca Central, Piso 05, Ciudad Universitaria de Caracas, Venezuela, correos electrónicos dir.dirección@gmail.com, sub.direccion@gmail.com, deu.contactos@gmail.com, Teléfonos: 605.3886/4940.`,
  CARPETA_PDF_ID: '',      // ID de carpeta en Drive para guardar PDFs (opcional)
  HISTORIAL_HABILITADO: true,
  ESTILOS: {
    FUENTE: 'Arial',
    TAMANO_NORMAL: 11,
    TAMANO_CODIGO: 14,
    MARGEN_SUPERIOR: 72,
    MARGEN_INFERIOR: 72,
    MARGEN_IZQUIERDO: 72,
    MARGEN_DERECHO: 72,
    INTERLINEADO: 1.15
  }
};

// ==================== FUNCIONES AUXILIARES ====================
const Utils = {
  zeroPad(num, length) {
    let s = String(num);
    while (s.length < length) s = '0' + s;
    return s;
  },

  fechaActualFormateada() {
    const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                   'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    const hoy = new Date();
    return `${hoy.getDate()} de ${meses[hoy.getMonth()]} de ${hoy.getFullYear()}`;
  },

  pedirDato(titulo, pregunta, obligatorio = true, defaultValue = '') {
    const ui = DocumentApp.getUi();
    for (let intentos = 0; intentos < 3; intentos++) {
      let texto = pregunta;
      if (defaultValue) texto += `\n\n(Valor por defecto: "${defaultValue}")`;
      const respuesta = ui.prompt(titulo, texto, ui.ButtonSet.OK_CANCEL);
      if (respuesta.getSelectedButton() !== ui.Button.OK) return null;
      let valor = respuesta.getResponseText().trim();
      if (valor === '' && defaultValue) valor = defaultValue;
      if (valor || !obligatorio) return valor || 'N/A';
      if (intentos < 2) ui.alert('⚠️ Campo Requerido', `Intento ${intentos+1} de 3`, ui.ButtonSet.OK);
    }
    ui.alert('❌ Demasiados intentos', 'Operación cancelada.', ui.ButtonSet.OK);
    return null;
  },

  mostrarError(mensaje, error = null) {
    console.error(mensaje, error);
    DocumentApp.getUi().alert('❌ Error', mensaje, DocumentApp.getUi().ButtonSet.OK);
  },

  mostrarExito(titulo, mensaje, url = null) {
    const ui = DocumentApp.getUi();
    let texto = mensaje;
    if (url) texto += `\n\nAbrir: ${url}`;
    ui.alert(`✅ ${titulo}`, texto, ui.ButtonSet.OK);
  }
};


// ==================== GESTOR DE PROPIEDADES ====================
const Props = {
  get(key, defaultValue = '1') {
    return PropertiesService.getScriptProperties().getProperty(key) || defaultValue;
  },
  set(key, value) {
    PropertiesService.getScriptProperties().setProperty(key, String(value));
  },
  incrementarContador() {
    const actual = parseInt(this.get('n_oficio', '1'));
    this.set('n_oficio', actual + 1);
    return actual;
  }
};

// ==================== GESTOR DE HISTORIAL ====================
const Historial = {
  getHoja() {
    const id = Props.get('HISTORIAL_ID', '');
    if (!id) return null;
    try {
      const ss = SpreadsheetApp.openById(id);
      let hoja = ss.getSheetByName('Registros');
      if (!hoja) {
        hoja = ss.insertSheet('Registros');
        hoja.appendRow(['Código', 'Director', 'Cargo', 'Escuela', 'Asunto', 'Fecha', 'Enlace PDF']);
      }
      return hoja;
    } catch (e) {
      Utils.mostrarError('No se pudo acceder al historial. Verifica el ID.', e);
      return null;
    }
  },

  agregarRegistro(datos) {
    if (!CONFIG.HISTORIAL_HABILITADO) return;
    const hoja = this.getHoja();
    if (!hoja) return;
    try {
      hoja.appendRow([
        datos.codigo,
        datos.director,
        datos.cargo,
        datos.escuela,
        datos.asunto,
        datos.fecha,
        datos.urlPDF
      ]);
    } catch (e) {
      Utils.mostrarError('Error al guardar en historial', e);
    }
  },

  configurar() {
    const ui = DocumentApp.getUi();
    const respuesta = ui.prompt(
      '📊 Configurar Historial',
      'ID de hoja de cálculo (vacío para crear una nueva):',
      ui.ButtonSet.OK_CANCEL
    );
    if (respuesta.getSelectedButton() !== ui.Button.OK) return;
    let id = respuesta.getResponseText().trim();
    try {
      if (!id) {
        const ss = SpreadsheetApp.create(`Historial DEU ${new Date().getFullYear()}`);
        id = ss.getId();
        const hoja = ss.getSheets()[0];
        hoja.setName('Registros');
        hoja.appendRow(['Código', 'Director', 'Cargo', 'Escuela', 'Asunto', 'Fecha', 'Enlace PDF']);
        Utils.mostrarExito('Historial Creado', `ID: ${id}\n\nURL: ${ss.getUrl()}`);
      } else {
        SpreadsheetApp.openById(id); // verifica que existe
        Utils.mostrarExito('Historial Configurado', 'Hoja vinculada correctamente.');
      }
      Props.set('HISTORIAL_ID', id);
    } catch (e) {
      Utils.mostrarError('No se pudo configurar el historial. Verifica el ID.', e);
    }
  }
};

// ==================== GESTOR DE DOCUMENTOS ====================
const Documento = {
  crear(nombreArchivo, datos) {
    const doc = DocumentApp.create(nombreArchivo);
    const body = doc.getBody();

    // Aplicar estilos generales al documento
    body.setFontSize(CONFIG.ESTILOS.TAMANO_NORMAL);
    body.setFontFamily(CONFIG.ESTILOS.FUENTE);
    body.setMarginTop(CONFIG.ESTILOS.MARGEN_SUPERIOR);
    body.setMarginBottom(CONFIG.ESTILOS.MARGEN_INFERIOR);
    body.setMarginLeft(CONFIG.ESTILOS.MARGEN_IZQUIERDO);
    body.setMarginRight(CONFIG.ESTILOS.MARGEN_DERECHO);
    body.setLineSpacing(CONFIG.ESTILOS.INTERLINEADO);

    // Insertar líneas de la plantilla con formato específico
    const lineas = CONFIG.PLANTILLA.split('\n');
    lineas.forEach(linea => {
      const parrafo = body.appendParagraph(linea);
      
      // Aplicar formato según el contenido de la línea
      if (linea.includes('{{codigo}}')) {
        parrafo.setBold(true).setFontSize(CONFIG.ESTILOS.TAMANO_CODIGO);
      } else if (linea.includes('Asunto:')) {
        parrafo.setBold(true);
      } else if (linea.includes('Atentamente.') || linea.includes('Profa. Mercy Ospina')) {
        parrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    });

    // Reemplazar etiquetas (el formato se conserva)
    const reemplazos = {
      '{{codigo}}': datos.codigo,
      '{{fecha}}': datos.fecha,
      '{{director}}': datos.director,
      '{{cargo}}': datos.cargo,
      '{{escuela}}': datos.escuela,
      '{{asunto}}': datos.asunto
    };
    Object.keys(reemplazos).forEach(etiqueta => {
      body.replaceText(etiqueta, reemplazos[etiqueta]);
    });

    doc.saveAndClose();
    return doc;
  },



  generarPDF(doc, nombreArchivo) {
    const pdfBlob = doc.getAs(MimeType.PDF).setName(`${nombreArchivo}.pdf`);
    let pdfFile;
    if (CONFIG.CARPETA_PDF_ID) {
      const carpeta = DriveApp.getFolderById(CONFIG.CARPETA_PDF_ID);
      pdfFile = carpeta.createFile(pdfBlob);
    } else {
      pdfFile = DriveApp.createFile(pdfBlob);
    }
    return pdfFile;
  },

  implementarPlantillaEnDocumentoActual() {
    const ui = DocumentApp.getUi();
    const respuesta = ui.alert(
      '⚠️ Reemplazar contenido',
      '¿Estás seguro de reemplazar TODO el contenido de este documento con la plantilla base?\n\nEsta acción no se puede deshacer.',
      ui.ButtonSet.YES_NO
    );
    if (respuesta !== ui.Button.YES) return;
    try {
      const doc = DocumentApp.getActiveDocument();
      doc.getBody().clear().setText(CONFIG.PLANTILLA);
      doc.saveAndClose();
      Utils.mostrarExito('Plantilla implementada', 'El documento ahora contiene la plantilla base.');
    } catch (e) {
      Utils.mostrarError('Error al implementar plantilla', e);
    }
  }
};

// ==================== MENÚ PRINCIPAL ====================
function onOpen() {
  try {
    const ui = DocumentApp.getUi();
    ui.createMenu('🏛️ SISTEMA DEU')
      .addSubMenu(ui.createMenu('⚙️ Configuración')
        .addItem('🔄 Resetear Contador', 'resetearContador')
        .addItem('👁️ Ver Estado', 'verEstadoSistema')
        .addItem('📊 Configurar Historial', 'configurarHistorial'))
      .addSeparator()
      .addItem('📄 Generar Oficio', 'procesarOficioCompleto')
      .addItem('📄 Crear documento con plantilla', 'crearPlantillaEjemplo')
      .addItem('📄 Mostrar plantilla en este documento', 'implementarPlantillaEnDocumento')
      .addSeparator()
      .addItem('❓ Ayuda', 'mostrarAyuda')
      .addToUi();
  } catch (e) {
    console.log('onOpen ejecutado en contexto sin UI');
  }
}

// ==================== FUNCIONES DEL MENÚ ====================
function verEstadoSistema() {
  const contador = Props.get('n_oficio', '1');
  const siguiente = Utils.zeroPad(parseInt(contador), 3);
  const historialId = Props.get('HISTORIAL_ID', '');
  const estadoHistorial = historialId ? '✅ Configurado' : '❌ No configurado';
  const carpetaPDF = CONFIG.CARPETA_PDF_ID ? '✅ Configurada' : '❌ No configurada (raíz)';

  const mensaje = `📊 ESTADO DEL SISTEMA

📌 Plantilla: Interna (test2.pdf) con formato profesional
🖼️ Logos: Integrados (Base64) - no requieren configuración
📁 Carpeta PDF: ${carpetaPDF}
🔢 Próximo número interno: ${siguiente}
📈 Historial: ${estadoHistorial}`;

  DocumentApp.getUi().alert('📊 Estado del Sistema', mensaje, DocumentApp.getUi().ButtonSet.OK);
}

function resetearContador() {
  const ui = DocumentApp.getUi();
  const actual = Props.get('n_oficio', '1');
  const respuesta = ui.alert(
    '⚠️ Resetear Contador',
    `¿Reiniciar a 001?\nNúmero actual: ${Utils.zeroPad(parseInt(actual), 3)}`,
    ui.ButtonSet.YES_NO
  );
  if (respuesta === ui.Button.YES) {
    Props.set('n_oficio', '1');
    ui.alert('✅ Contador Reseteado', 'Próximo número interno: 001', ui.ButtonSet.OK);
  }
}

function configurarHistorial() {
  Historial.configurar();
}

function mostrarAyuda() {
  const html = `<!DOCTYPE html>
<html>
<head><style>
  body { font-family: 'Segoe UI', Arial, sans-serif; padding:20px; background:#f5f5f5; }
  .container { background:white; border-radius:10px; padding:20px; box-shadow:0 2px 10px rgba(0,0,0,0.1); }
  h2 { color:#1a73e8; border-bottom:2px solid #1a73e8; padding-bottom:10px; }
  .etiqueta { background:#e8f0fe; color:#1a73e8; font-family:monospace; padding:4px 8px; border-radius:4px; display:inline-block; margin:2px 0; }
  .codigo { background:#f0f0f0; padding:10px; border-radius:5px; font-size:12px; }
</style></head>
<body>
<div class="container">
  <h2>📘 Manual de Etiquetas</h2>
  <p><span class="etiqueta">{{codigo}}</span> - Código (ej. DEU-GSU)</p>
  <p><span class="etiqueta">{{fecha}}</span> - Fecha (ej. 27 de febrero de 2026)</p>
  <p><span class="etiqueta">{{director}}</span> - Nombre del director</p>
  <p><span class="etiqueta">{{cargo}}</span> - Cargo del destinatario</p>
  <p><span class="etiqueta">{{escuela}}</span> - Escuela o facultad</p>
  <p><span class="etiqueta">{{asunto}}</span> - Asunto del oficio</p>
  <h3>💡 Plantilla actual:</h3>
  <div class="codigo">${CONFIG.PLANTILLA.replace(/\n/g, '<br>')}</div>
  <p style="text-align:center; margin-top:20px;">Las etiquetas DEBEN ir exactamente como se muestran.</p>
</div>
</body>
</html>`;
  DocumentApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(450).setHeight(550),
    '❓ Ayuda - Sistema DEU'
  );
}

function crearPlantillaEjemplo() {
  try {
    const doc = DocumentApp.create('Plantilla DEU - Test2');
    doc.getBody().setText(CONFIG.PLANTILLA);
    doc.saveAndClose();
    Utils.mostrarExito('Plantilla creada', `Documento: ${doc.getUrl()}`);
  } catch (e) {
    Utils.mostrarError('No se pudo crear la plantilla.', e);
  }
}

function implementarPlantillaEnDocumento() {
  Documento.implementarPlantillaEnDocumentoActual();
}

// ==================== PROCESO PRINCIPAL ====================
function procesarOficioCompleto() {
  const ui = DocumentApp.getUi();

  // 1. Recolectar datos
  const datos = {};
  datos.codigo = Utils.pedirDato('🔖 Código', 'Código del oficio:', true, 'DEU-GSU');
  if (!datos.codigo) return;

  datos.fecha = Utils.pedirDato('📅 Fecha', 'Fecha:', true, Utils.fechaActualFormateada());
  if (!datos.fecha) return;

  datos.director = Utils.pedirDato('👤 Director', 'Nombre del director:');
  if (!datos.director) return;

  datos.cargo = Utils.pedirDato('👔 Cargo', 'Cargo del destinatario:');
  if (!datos.cargo) return;

  datos.escuela = Utils.pedirDato('🏛️ Escuela', 'Nombre de la escuela:');
  if (!datos.escuela) return;

  datos.asunto = Utils.pedirDato('📋 Asunto', 'Asunto del oficio:');
  if (!datos.asunto) return;

  // 2. Generar documentos
  try {
    const contador = Props.incrementarContador();
    const numero = Utils.zeroPad(contador, 3);
    const anio = new Date().getFullYear();
    const nombreBase = datos.escuela.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    const nombreArchivo = `Oficio_${numero}_${anio}_${nombreBase}`;

    const doc = Documento.crear(nombreArchivo, datos);
    const pdfFile = Documento.generarPDF(doc, nombreArchivo);

    // 3. Guardar en historial
    Historial.agregarRegistro({
      ...datos,
      urlPDF: pdfFile.getUrl()
    });

    // 4. Mostrar resultado
    const html = `<!DOCTYPE html>
<html>
<head><style>
  body { font-family:'Segoe UI',Arial,sans-serif; text-align:center; padding:20px; background:#f0f7ff; }
  .card { background:#fff; padding:20px; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.1); border-left:5px solid #1a73e8; }
  h2 { color:#1a73e8; margin-top:0; }
  .info { text-align:left; background:#f8f9fa; padding:15px; border-radius:8px; margin:15px 0; }
  .info b { color:#1a73e8; width:100px; display:inline-block; }
  .btn { background:#1a73e8; color:#fff; padding:12px 25px; border:none; border-radius:25px; cursor:pointer; font-size:16px; font-weight:bold; margin:10px 5px; text-decoration:none; display:inline-block; }
  .btn.secondary { background:#34a853; }
  .warning { color:#666; font-size:11px; margin-top:15px; }
</style></head>
<body>
<div class="card">
  <h2>✅ ¡Oficio Generado!</h2>
  <div class="info">
    <p><b>Código:</b> ${datos.codigo}</p>
    <p><b>Director:</b> ${datos.director}</p>
    <p><b>Escuela:</b> ${datos.escuela}</p>
    <p><b>Asunto:</b> ${datos.asunto}</p>
  </div>
  <a href="${pdfFile.getUrl()}" target="_blank" class="btn">📄 ABRIR PDF</a>
  <a href="${doc.getUrl()}" target="_blank" class="btn secondary">📝 EDITAR DOC</a>
  <p class="warning">Número interno: ${numero} (solo para control)</p>
</div>
</body>
</html>`;

    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(500), '🎉 Éxito');

  } catch (error) {
    Utils.mostrarError('Error al generar el oficio', error);
  }
}