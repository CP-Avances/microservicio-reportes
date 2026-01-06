package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ResumenAsistencia.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Service
public class ReporteAsistenciaService {

    public byte[] generarReporteResumenAsistenciaPDF(ReporteAsistenciaRequest request) {

        // Constantes DRY
        final float[] WIDTHS_6 = { 2, 2, 2, 2, 2, 2 };
        final float[] WIDTHS_2 = { 8, 2 };
        final float[] WIDTHS_16 = {
                0.7f, 1.6f,
                1.8f, 1.5f, 1.8f, 1.5f,
                1.8f, 1.5f, 1.8f, 1.5f,
                1.5f, 1.5f, 2f, 1.7f,
                2.1f, 3f
        };

        final Color COLOR_FALTA_TIMBRE = new Color(0xEE4444);
        final Color COLOR_ATRASO = new Color(0xEEE344);
        final Color COLOR_SALIDA_ANTICIPADA = new Color(0x4499EE);
        final Color COLOR_EXCESO_ALIMENTACION = new Color(0x55EE44);
        final Color COLOR_VACACIONES = new Color(0xE68A2E);

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // Inicialización de recursos
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // Logo y encabezados (adaptable: si faltan datos, se omiten)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            String titulo = "RESUMEN DE ASISTENCIA - "
                    + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Tabla de códigos de color
            PdfPTable colores = new PdfPTable(6);
            colores.setWidthPercentage(100);
            colores.setWidths(WIDTHS_6);
            colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", COLOR_FALTA_TIMBRE));
            colores.addCell(ReporteUtil.celdaEncabezado("ATRASO", COLOR_ATRASO));
            colores.addCell(ReporteUtil.celdaEncabezado("SALIDA ANTICIPADA", COLOR_SALIDA_ANTICIPADA));
            colores.addCell(ReporteUtil.celdaEncabezado("EXCESO DE ALIMENTACIÓN",
                    COLOR_EXCESO_ALIMENTACION));
            colores.addCell(ReporteUtil.celdaEncabezado("VACACIONES", COLOR_VACACIONES));
            colores.setSpacingAfter(10f);
            document.add(colores);

            // Contador global
            AtomicInteger totalRegistros = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados().forEach(
                            emp -> totalRegistros.addAndGet(emp.getTLaborado().size())));

            // Tabla N° registros
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(WIDTHS_2);
            tituloTabla.setSpacingAfter(5f);

            PdfPCell celdaTitulo = new PdfPCell(
                    new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + totalRegistros.get(),
                            ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            tituloTabla.setSpacingAfter(10f);
            document.add(tituloTabla);

            // Grupos / Empleados
            for (GrupoAsistenciaDTO grupo : request.getGrupos()) {
                for (EmpleadoAsistenciaDTO emp : grupo.getEmpleados()) {
                    double totalAtrasos = 0;
                    double totalSalidasAnticipadas = 0;
                    double totalAlimentacionTomado = 0;
                    double totalAlimentacionAsignado = 0;
                    double totalLaborado = 0;

                    // Información del empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",
                            emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                            emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(),
                            zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:",
                            emp.getRegimen(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:",
                            emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(),
                            zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    tablaContenedora.setSpacingAfter(5f);
                    document.add(tablaContenedora);

                    // Encabezado (16 columnas)
                    PdfPTable encabezado = new PdfPTable(16);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(WIDTHS_16);

                    encabezado.addCell(ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 2, 1));
                    encabezado.addCell(ReporteUtil.crearCelda("FECHA",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                    encabezado.addCell(ReporteUtil.crearCelda("ENTRADA",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("INICIO ALIMENTACIÓN",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("FIN ALIMENTACIÓN",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("SALIDA",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));

                    encabezado.addCell(ReporteUtil.crearCelda("ATRASO",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(ReporteUtil.crearCelda("SALIDA ANTICIPADA",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                    encabezado.addCell(ReporteUtil.crearCelda("T. ALIMENTACIÓN",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("TIEMPO LABORADO",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(ReporteUtil.crearCelda("OBSERVACIONES",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                    for (int i = 0; i < 4; i++) {
                        encabezado.addCell(ReporteUtil.crearCelda("HORARIO",
                                ReporteUtil.fuenteEncabezado(), colorPrincipal));
                        encabezado.addCell(ReporteUtil.crearCelda("TIMBRE",
                                ReporteUtil.fuenteEncabezado(), colorSecundario));
                    }
                    encabezado.addCell(ReporteUtil.crearCelda("ASIGNADO",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    encabezado.addCell(ReporteUtil.crearCelda("TOMADO",
                            ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    encabezado.setSpacingAfter(0f);
                    document.add(encabezado);

                    // Tabla de datos
                    PdfPTable tablaData = new PdfPTable(16);
                    tablaData.setWidthPercentage(100);
                    tablaData.setWidths(WIDTHS_16);

                    int contador = 1;
                    for (RegistroAsistenciaDTO reg : emp.getTLaborado()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                        tablaData.addCell(ReporteUtil.crearCelda(String.valueOf(contador),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(
                                ReporteUtil.formatearFechaConDia(
                                        reg.getEntrada().getFecha_horario()),
                                ReporteUtil.fuenteTexto(), fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                extraerHora(reg.getEntrada().getFecha_hora_horario()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(
                                formatearTimbre(reg.getEntrada()
                                        .getFecha_hora_horario(),
                                        reg.getEntrada().getFecha_hora_timbre()),
                                ReporteUtil.fuenteTexto(),
                                getColorTimbre(
                                        formatearTimbre(reg.getEntrada()
                                                .getFecha_hora_horario(),
                                                reg.getEntrada().getFecha_hora_timbre()),
                                        fondo, COLOR_FALTA_TIMBRE)));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                extraerHora(reg.getInicioAlimentacion()
                                        .getFecha_hora_horario()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(
                                formatearTimbre(reg.getInicioAlimentacion()
                                        .getFecha_hora_horario(),
                                        reg.getInicioAlimentacion()
                                                .getFecha_hora_timbre()),
                                ReporteUtil.fuenteTexto(),
                                getColorTimbre(
                                        formatearTimbre(reg
                                                .getInicioAlimentacion()
                                                .getFecha_hora_horario(),
                                                reg.getInicioAlimentacion()
                                                        .getFecha_hora_timbre()),
                                        fondo, COLOR_FALTA_TIMBRE)));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                extraerHora(reg.getFinAlimentacion()
                                        .getFecha_hora_horario()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(
                                formatearTimbre(reg.getFinAlimentacion()
                                        .getFecha_hora_horario(),
                                        reg.getFinAlimentacion()
                                                .getFecha_hora_timbre()),
                                ReporteUtil.fuenteTexto(),
                                getColorTimbre(
                                        formatearTimbre(reg.getFinAlimentacion()
                                                .getFecha_hora_horario(),
                                                reg.getFinAlimentacion()
                                                        .getFecha_hora_timbre()),
                                        fondo, COLOR_FALTA_TIMBRE)));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                extraerHora(reg.getSalida().getFecha_hora_horario()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(
                                formatearTimbre(reg.getSalida().getFecha_hora_horario(),
                                        reg.getSalida().getFecha_hora_timbre()),
                                ReporteUtil.fuenteTexto(),
                                getColorTimbre(
                                        formatearTimbre(reg.getSalida()
                                                .getFecha_hora_horario(),
                                                reg.getSalida().getFecha_hora_timbre()),
                                        fondo, COLOR_FALTA_TIMBRE)));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getMinAtrasos()),
                                ReporteUtil.fuenteTexto(),
                                reg.getMinAtrasos() != null && reg.getMinAtrasos() > 0
                                        ? COLOR_ATRASO
                                        : fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getMinSalidasAnticipadas()),
                                ReporteUtil.fuenteTexto(),
                                reg.getMinSalidasAnticipadas() != null
                                        && reg.getMinSalidasAnticipadas() > 0
                                                ? COLOR_SALIDA_ANTICIPADA
                                                : fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getInicioAlimentacion()
                                        .getMinutos_alimentacion()),
                                ReporteUtil.fuenteTexto(), fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getMinAlimentacion()),
                                ReporteUtil.fuenteTexto(),
                                (reg.getInicioAlimentacion()
                                        .getMinutos_alimentacion() != null &&
                                        reg.getMinAlimentacion() != null &&
                                        reg.getMinAlimentacion() > reg
                                                .getInicioAlimentacion()
                                                .getMinutos_alimentacion())
                                                        ? COLOR_EXCESO_ALIMENTACION
                                                        : fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getMinLaborados()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(),
                                fondo));

                        // Acumuladores
                        totalAtrasos += reg.getMinAtrasos() != null ? reg.getMinAtrasos() : 0;
                        totalSalidasAnticipadas += reg.getMinSalidasAnticipadas() != null
                                ? reg.getMinSalidasAnticipadas()
                                : 0;
                        totalAlimentacionTomado += reg.getMinAlimentacion() != null
                                ? reg.getMinAlimentacion()
                                : 0;
                        totalAlimentacionAsignado += reg.getInicioAlimentacion()
                                .getMinutos_alimentacion() != null ? reg
                                        .getInicioAlimentacion()
                                        .getMinutos_alimentacion() : 0;
                        totalLaborado += reg.getMinLaborados() != null ? reg.getMinLaborados()
                                : 0;

                        contador++;
                    }

                    // Fila final de totales (vacías hasta la columna 10)
                    for (int i = 0; i < 9; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("",
                                ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER);
                        tablaData.addCell(celdaVacia);
                    }
                    tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(),
                            Color.WHITE));

                    tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalAtrasos),
                            ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaData.addCell(ReporteUtil.crearCelda(
                            convertirMinutosATiempo(totalSalidasAnticipadas),
                            ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaData.addCell(ReporteUtil.crearCelda(
                            convertirMinutosATiempo(totalAlimentacionAsignado),
                            ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaData.addCell(ReporteUtil.crearCelda(
                            convertirMinutosATiempo(totalAlimentacionTomado),
                            ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalLaborado),
                            ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(),
                            Color.WHITE));

                    tablaData.setSpacingAfter(10f);
                    document.add(tablaData);
                }
            }

            // Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper tuyo lanza IAEx, dejamos que el controller decida 400 si
            // corresponde
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar ResumenAsistencia.pdf", e);
        } finally {
            // Ciclo de recursos garantizado
            if (document != null && document.isOpen()) {
                try {
                    document.close();
                } catch (Exception ignore) {
                }
            }
            if (writer != null) {
                try {
                    writer.close();
                } catch (Exception ignore) {
                }
            }
            if (baos != null) {
                try {
                    baos.close();
                } catch (Exception ignore) {
                }
            }
        }
    }

    // ===== XLSX (nuevo) =====
    public byte[] generarReporteResumenAsistenciaXLSX(ReporteAsistenciaRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Resumen_asistencia";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:W5 (B=1 .. W=22 en 0-based)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 22;

        final String[] HEADERS = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL", "RÉGIMEN",
                "DEPARTAMENTO", "CARGO", "FECHA", "HORARIO ENTRADA", "TIMBRE ENTRADA",
                "HORARIO INICIO ALIMENTACIÓN", "TIMBRE INICIO ALIMENTACIÓN",
                "HORARIO FIN ALIMENTACIÓN", "TIMBRE FIN ALIMENTACIÓN",
                "HORARIO SALIDA", "TIMBRE SALIDA",
                "ATRASO", "SALIDA ANTICIPADA",
                "TIEMPO ALIMENTACIÓN ASIGNADO", "TIEMPO ALIMENTACIÓN HH:MM:SS",
                "TIEMPO LABORADO HH:MM:SS"
        };

        final int[] ANCHOS = {
                10, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 40,
                40, 40, 40
        };

        // ===========================================================
        // 1) Inicialización de recursos (try-with-resources obligatorio)
        // ===========================================================
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);

            // (Opcional) Congelar encabezado para mejor UX
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // =====================================
            // 2) Construcción (helpers ya existentes)
            // =====================================

            // 2.1 Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2.2 Merges B1:W5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 2.3 Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "RESUMEN DE ASISTENCIA", estiloTitulo);
            String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                    + safe(request.getFechaFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            // 2.4 Encabezados + anchos
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 2.5 Cuerpo (aplanado grupos → empleados → tLaborado)
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getGrupos() != null) {
                for (GrupoAsistenciaDTO grupo : request.getGrupos()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoAsistenciaDTO usu : grupo.getEmpleados()) {
                        String apenom = (safe(usu.getApellido()) + " " + safe(usu.getNombre())).trim();
                        if (usu.getTLaborado() == null)
                            continue;

                        for (RegistroAsistenciaDTO t : usu.getTLaborado()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int col = 0;

                            // Cálculos (replica TS)
                            String entradaHorario = hora(
                                    horaDe(t.getEntrada() != null ? t.getEntrada().getFecha_hora_horario() : null));
                            String salidaHorario = hora(
                                    horaDe(t.getSalida() != null ? t.getSalida().getFecha_hora_horario() : null));
                            String iniAliHorario = "EAS".equals(safe(t.getTipo()))
                                    ? hora(horaDe(t.getInicioAlimentacion() != null
                                            ? t.getInicioAlimentacion().getFecha_hora_horario()
                                            : null))
                                    : "";
                            String finAliHorario = "EAS".equals(safe(t.getTipo()))
                                    ? hora(horaDe(t.getFinAlimentacion() != null
                                            ? t.getFinAlimentacion().getFecha_hora_horario()
                                            : null))
                                    : "";

                            boolean control = t.getControl();

                            String entrada = timbreValor(
                                    t.getEntrada() != null ? t.getEntrada().getFecha_hora_horario() : null,
                                    t.getEntrada() != null ? t.getEntrada().getFecha_hora_timbre() : null,
                                    safe(t.getOrigen()), control);

                            String salida = timbreValor(
                                    t.getSalida() != null ? t.getSalida().getFecha_hora_horario() : null,
                                    t.getSalida() != null ? t.getSalida().getFecha_hora_timbre() : null,
                                    safe(t.getOrigen()), control);

                            String iniAli = "EAS".equals(safe(t.getTipo()))
                                    ? timbreValor(
                                            t.getInicioAlimentacion() != null
                                                    ? t.getInicioAlimentacion().getFecha_hora_horario()
                                                    : null,
                                            t.getInicioAlimentacion() != null
                                                    ? t.getInicioAlimentacion().getFecha_hora_timbre()
                                                    : null,
                                            safe(t.getOrigen()), control)
                                    : "";

                            String finAli = "EAS".equals(safe(t.getTipo()))
                                    ? timbreValor(
                                            t.getFinAlimentacion() != null
                                                    ? t.getFinAlimentacion().getFecha_hora_horario()
                                                    : null,
                                            t.getFinAlimentacion() != null
                                                    ? t.getFinAlimentacion().getFecha_hora_timbre()
                                                    : null,
                                            safe(t.getOrigen()), control)
                                    : "";

                            Double asignMin = "EAS".equals(safe(t.getTipo())) && t.getInicioAlimentacion() != null
                                    ? t.getInicioAlimentacion().getMinutos_alimentacion()
                                    : 0d;
                            String alimentacionAsignada = convertirMinutosATiempo(asignMin);

                            String tiempoAlimentacion = convertirMinutosATiempo(t.getMinAlimentacion());
                            Double minsLaborados = control ? t.getMinLaborados() : t.getMinPlanificados();
                            String tiempoLaborado = convertirMinutosATiempo(minsLaborados);

                            String tiempoAtraso = convertirMinutosATiempo(t.getMinAtrasos());
                            String tiempoSalidaAnt = convertirMinutosATiempo(t.getMinSalidasAnticipadas());

                            // Escritura de fila
                            UtilExcel.establecerValor(r, col++, item++, null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getCodigo()), null);
                            UtilExcel.establecerTexto(r, col++, apenom, null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getCiudad()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getSucursal()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getRegimen()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getCargo()), null);

                            UtilExcel.establecerTexto(r, col++,
                                    safe(t.getEntrada() != null ? t.getEntrada().getFecha_hora_horario() : null), null); // FECHA
                                                                                                                         // (ISO
                                                                                                                         // completa)
                            UtilExcel.establecerTexto(r, col++, entradaHorario, null);
                            UtilExcel.establecerTexto(r, col++, entrada, null);
                            UtilExcel.establecerTexto(r, col++, iniAliHorario, null);
                            UtilExcel.establecerTexto(r, col++, iniAli, null);
                            UtilExcel.establecerTexto(r, col++, finAliHorario, null);
                            UtilExcel.establecerTexto(r, col++, finAli, null);
                            UtilExcel.establecerTexto(r, col++, salidaHorario, null);
                            UtilExcel.establecerTexto(r, col++, salida, null);

                            UtilExcel.establecerTexto(r, col++, tiempoAtraso, null);
                            UtilExcel.establecerTexto(r, col++, tiempoSalidaAnt, null);
                            UtilExcel.establecerTexto(r, col++, alimentacionAsignada, null);
                            UtilExcel.establecerTexto(r, col++, tiempoAlimentacion, null);
                            UtilExcel.establecerTexto(r, col++, tiempoLaborado, null);
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 2.6 Estilos de cuerpo
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde,
                        true);
            }

            // 2.7 Tabla estilizada + filtros (ITEM sin filtro)
            if (ultimaFila >= filaDatosIni) {
                boolean[] filtros = new boolean[HEADERS.length];
                for (int i = 0; i < filtros.length; i++)
                    filtros[i] = true;
                filtros[0] = false;

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "ResumenGeneralReporteTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        filtros);
            }

            // ======================
            // 3) Cierre y retorno
            // ======================
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación/entrada inválida (dejar que el controller lo mapee a 400 si
            // aplica)
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ResumenAsistencia.xlsx", e);
        }
    }

    // ===== Utilidades privadas =====
    private String hora(String hhmmss) {
        return (hhmmss == null) ? "" : hhmmss;
    }

    private String horaDe(String fechaHora) {
        if (fechaHora == null || !fechaHora.contains(" "))
            return "";
        String[] p = fechaHora.split(" ");
        return (p.length > 1) ? p[1] : "";
    }

    private String timbreValor(String horario, String timbre, String origen, Boolean control) {
        String hhHorario = horaDe(horario);
        if (hhHorario == null || hhHorario.isEmpty())
            return "";
        if (timbre != null && !timbre.trim().isEmpty()) {
            return horaDe(timbre);
        }
        if ("L".equals(origen) || "FD".equals(origen))
            return origen;
        return (control != null && control) ? "FT" : "SCA";
    }

    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String convertirMinutosATiempo(Double minutos) {
        if (minutos == null || minutos <= 0)
            return "00:00:00";
        int totalSegundos = (int) Math.round(minutos * 60);
        int horas = totalSegundos / 3600;
        int mins = (totalSegundos % 3600) / 60;
        int segundos = totalSegundos % 60;
        return String.format("%02d:%02d:%02d", horas, mins, segundos);
    }

    private String extraerHora(String fechaHora) {
        if (fechaHora == null || !fechaHora.contains(" "))
            return "";
        return fechaHora.split(" ")[1];
    }

    private String formatearTimbre(String horario, String timbre) {
        String horaHorario = extraerHora(horario);

        if (horaHorario == null || horaHorario.isEmpty()) {
            return "";
        }

        if (timbre == null || timbre.trim().isEmpty()) {
            if ("00:00:00".equals(horaHorario) || "23:59:00".equals(horaHorario)) {
                return "L";
            }
            return "FT";
        }

        return extraerHora(timbre);
    }

    private Color getColorTimbre(String valor, Color porDefecto, Color colorFT) {
        if ("FT".equals(valor))
            return colorFT;
        return porDefecto;
    }

}
