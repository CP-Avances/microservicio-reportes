package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.ReporteSalidasAnticipadasRequest;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.GrupoSalidasDTO;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.EmpleadoSalidaDTO;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.SalidaDTO;
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
public class ReporteSalidasAnticipadasService {

    public byte[] generarReportePDF(
            ReporteSalidasAnticipadasRequest request) {

        final float[] WIDTHS_TITULO_TABLA = { 8f, 2f };
        final float[] WIDTHS_INFO_EMPLEADO = { 4f, 4f, 4f };

        final float[] WIDTHS_SALIDAS = {
                1f, // N°
                1.8f, // Horario fecha
                2f, // Horario hora
                2f, // Timbre fecha
                2f, // Timbre hora
                2.5f, // Tipo permiso
                1.8f, // Desde
                1.8f, // Hasta
                1.4f, // Permiso tiempo
                1.2f, // Permiso decimal
                1.8f, // Salida anticipada tiempo
                1.4f // Salida anticipada decimal
        };

        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_AFTER_TITULO_TABLA = 10f;
        final float SPACING_BEFORE_SALIDAS = 5f;
        final float PADDING_ENCABEZADOS = 5f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(
                request.getColorPrincipal());

        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(
                request.getColorSecundario());

        final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();

        final String TITULO = "SALIDAS ANTICIPADAS - "
                + (request.getOpcionBusqueda() != null
                        && request.getOpcionBusqueda() == 1
                                ? "ACTIVOS"
                                : "INACTIVOS");

        final String PERIODO = "PERIODO DEL: "
                + safe(request.getFechaInicio())
                + " AL "
                + safe(request.getFechaFin());

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            /*
             * 1. INICIALIZACIÓN DEL DOCUMENTO
             */
            baos = new ByteArrayOutputStream();

            document = new Document(
                    PageSize.A4.rotate(),
                    30,
                    30,
                    30,
                    50
            );

            writer = PdfWriter.getInstance(
                    document,
                    baos);

            writer.setPageEvent(
                    new ConfiguracionPaginaPDF(
                            request.getUsuario(),
                            request.getFraseMarcaAgua(),
                            request.getColorPrincipal()));

            document.open();

            /*
             * 2. LOGO Y TÍTULOS
             */
            Image logo = ReporteUtil.obtenerLogo(
                    request.getLogoBase64());

            if (logo != null) {
                document.add(logo);
            }

            document.add(
                    ReporteUtil.crearTituloEmpresa(
                            request.getEmpresa()));

            document.add(
                    ReporteUtil.crearTituloReporte(
                            TITULO));

            document.add(
                    ReporteUtil.crearTituloPeriodo(
                            PERIODO));

            /*
             * 3. TOTAL GENERAL DE REGISTROS
             */
            AtomicInteger contadorGlobal = new AtomicInteger();

            if (request.getGrupos() != null) {
                request.getGrupos().forEach(grupo -> {
                    if (grupo.getEmpleados() == null) {
                        return;
                    }

                    grupo.getEmpleados().forEach(empleado -> {
                        int cantidadSalidas = empleado.getSalidas() != null
                                ? empleado.getSalidas().size()
                                : 0;

                        contadorGlobal.addAndGet(
                                cantidadSalidas);
                    });
                });
            }

            /*
             * 4. CABECERA LISTA EMPLEADOS
             */
            PdfPTable tituloTabla = new PdfPTable(2);

            tituloTabla.setWidthPercentage(
                    WIDTH_PERCENT_100);

            tituloTabla.setWidths(
                    WIDTHS_TITULO_TABLA);

            tituloTabla.setSpacingAfter(
                    SPACING_AFTER_TITULO_TABLA);

            PdfPCell celdaTitulo = new PdfPCell(
                    new Phrase(
                            "LISTA EMPLEADOS",
                            ReporteUtil.fuenteEncabezadoTablaData()));

            celdaTitulo.setBackgroundColor(
                    COLOR_SECUNDARIO);

            celdaTitulo.setPadding(
                    PADDING_ENCABEZADOS);

            celdaTitulo.setBorder(
                    Rectangle.TOP
                            | Rectangle.BOTTOM
                            | Rectangle.LEFT);

            tituloTabla.addCell(
                    celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase(
                            "N° Registros: "
                                    + contadorGlobal.get(),
                            ReporteUtil.fuenteEncabezadoTablaData()));

            celdaContador.setBackgroundColor(
                    COLOR_SECUNDARIO);

            celdaContador.setHorizontalAlignment(
                    Element.ALIGN_RIGHT);

            celdaContador.setVerticalAlignment(
                    Element.ALIGN_MIDDLE);

            celdaContador.setPadding(
                    PADDING_ENCABEZADOS);

            celdaContador.setBorder(
                    Rectangle.TOP
                            | Rectangle.BOTTOM
                            | Rectangle.RIGHT);

            tituloTabla.addCell(
                    celdaContador);

            document.add(
                    tituloTabla);

            int contador = 1;

            /*
             * 5. RECORRIDO DE GRUPOS Y EMPLEADOS
             */
            if (request.getGrupos() != null) {

                for (GrupoSalidasDTO grupo : request.getGrupos()) {

                    if (grupo.getEmpleados() == null) {
                        continue;
                    }

                    for (EmpleadoSalidaDTO empleado : grupo.getEmpleados()) {

                        /*
                         * INFORMACIÓN DEL EMPLEADO
                         */
                        PdfPTable infoEmpleado = new PdfPTable(3);

                        infoEmpleado.setWidthPercentage(
                                WIDTH_PERCENT_100);

                        infoEmpleado.setWidths(
                                WIDTHS_INFO_EMPLEADO);

                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta(
                                        "C.C.:",
                                        safe(empleado.getIdentificacion()),
                                        COLOR_ZEBRA));

                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta(
                                        "EMPLEADO:",
                                        (safe(empleado.getApellido())
                                                + " "
                                                + safe(empleado.getNombre())).trim(),
                                        COLOR_ZEBRA));

                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta(
                                        "COD:",
                                        safe(empleado.getCodigo()),
                                        COLOR_ZEBRA));

                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta(
                                        "RÉGIMEN LABORAL:",
                                        safe(empleado.getRegimen()),
                                        COLOR_ZEBRA));

                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta(
                                        "DEPARTAMENTO:",
                                        safe(empleado.getDepartamento()),
                                        COLOR_ZEBRA));

                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta(
                                        "CARGO:",
                                        safe(empleado.getCargo()),
                                        COLOR_ZEBRA));

                        PdfPTable tablaContenedora = new PdfPTable(1);

                        tablaContenedora.setWidthPercentage(
                                WIDTH_PERCENT_100);

                        PdfPCell contenedor = new PdfPCell(
                                infoEmpleado);

                        contenedor.setPadding(0);
                        contenedor.setBorder(
                                Rectangle.BOX);

                        tablaContenedora.addCell(
                                contenedor);

                        document.add(
                                tablaContenedora);

                        /*
                         * TABLA DE SALIDAS ANTICIPADAS
                         */
                        PdfPTable tablaSalidas = new PdfPTable(12);

                        tablaSalidas.setWidthPercentage(
                                WIDTH_PERCENT_100);

                        tablaSalidas.setSpacingBefore(
                                SPACING_BEFORE_SALIDAS);

                        tablaSalidas.setWidths(
                                WIDTHS_SALIDAS);

                        /*
                         * ENCABEZADO: PRIMERA FILA
                         */
                        PdfPCell encabezadoNumero = ReporteUtil.crearCelda(
                                "N°",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoNumero.setRowspan(2);
                        tablaSalidas.addCell(
                                encabezadoNumero);

                        PdfPCell encabezadoHorario = ReporteUtil.crearCelda(
                                "HORARIO",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoHorario.setColspan(2);
                        tablaSalidas.addCell(
                                encabezadoHorario);

                        PdfPCell encabezadoTimbre = ReporteUtil.crearCelda(
                                "TIMBRE",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoTimbre.setColspan(2);
                        tablaSalidas.addCell(
                                encabezadoTimbre);

                        PdfPCell encabezadoTipoPermiso = ReporteUtil.crearCelda(
                                "TIPO PERMISO",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoTipoPermiso.setRowspan(2);
                        tablaSalidas.addCell(
                                encabezadoTipoPermiso);

                        PdfPCell encabezadoDesde = ReporteUtil.crearCelda(
                                "DESDE",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoDesde.setRowspan(2);
                        tablaSalidas.addCell(
                                encabezadoDesde);

                        PdfPCell encabezadoHasta = ReporteUtil.crearCelda(
                                "HASTA",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoHasta.setRowspan(2);
                        tablaSalidas.addCell(
                                encabezadoHasta);

                        PdfPCell encabezadoPermiso = ReporteUtil.crearCelda(
                                "PERMISO",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoPermiso.setColspan(2);
                        tablaSalidas.addCell(
                                encabezadoPermiso);

                        PdfPCell encabezadoSalida = ReporteUtil.crearCelda(
                                "SALIDA ANTICIPADA",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);

                        encabezadoSalida.setColspan(2);
                        tablaSalidas.addCell(
                                encabezadoSalida);

                        /*
                         * ENCABEZADO: SEGUNDA FILA
                         */
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "FECHA",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "HORA",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "FECHA",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "HORA",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "TIEMPO",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "DECIMAL",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "TIEMPO",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "DECIMAL",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_PRIMARIO));

                        /*
                         * TOTALES DEL EMPLEADO
                         */
                        long totalSegundosPermiso = 0L;
                        double totalMinutosPermiso = 0d;

                        long totalSegundosSalida = 0L;
                        double totalMinutosSalida = 0d;

                        /*
                         * FILAS DE SALIDAS
                         */
                        if (empleado.getSalidas() != null) {

                            for (SalidaDTO salida : empleado.getSalidas()) {

                                Color fondo = contador % 2 == 0
                                        ? COLOR_ZEBRA
                                        : Color.WHITE;

                                String[] horario = splitFechaHora(
                                        salida.getFecha_hora_horario());

                                String[] timbre = splitFechaHora(
                                        salida.getFecha_hora_timbre());

                                /*
                                 * SALIDA ANTICIPADA RESTANTE
                                 */
                                long segundosSalida = salida.getDiferencia() != null
                                        ? Math.max(
                                                Math.round(
                                                        salida.getDiferencia()),
                                                0L)
                                        : 0L;

                                double minutosSalida = segundosSalida / 60.0;

                                String tiempoSalida = convertirMinutosATiempo(
                                        minutosSalida);

                                /*
                                 * DURACIÓN COMPLETA DEL PERMISO
                                 */
                                long segundosPermiso = salida.getDiferencia_permiso() != null
                                        ? Math.max(
                                                Math.round(
                                                        salida.getDiferencia_permiso()),
                                                0L)
                                        : 0L;

                                double minutosPermiso = segundosPermiso / 60.0;

                                boolean tienePermiso = segundosPermiso > 0
                                        || !safe(
                                                salida.getTipo_permiso()).isEmpty();

                                String tiempoPermiso = tienePermiso
                                        ? convertirMinutosATiempo(
                                                minutosPermiso)
                                        : "";

                                String decimalPermiso = tienePermiso
                                        ? String.format(
                                                "%.2f",
                                                minutosPermiso)
                                        : "";

                                /*
                                 * NÚMERO
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                String.valueOf(contador),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * HORARIO
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                ReporteUtil.formatearFechaConDia(
                                                        horario[0]),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                horario[1],
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * TIMBRE
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                ReporteUtil.formatearFechaConDia(
                                                        timbre[0]),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                timbre[1],
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * TIPO DE PERMISO
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                safe(
                                                        salida.getTipo_permiso()),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * DESDE
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                formatearFechaHoraReporte(
                                                        salida.getDesde()),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * HASTA
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                formatearFechaHoraReporte(
                                                        salida.getHasta()),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * PERMISO: TIEMPO
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                tiempoPermiso,
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * PERMISO: DECIMAL
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                decimalPermiso,
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * SALIDA ANTICIPADA: TIEMPO
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                tiempoSalida,
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * SALIDA ANTICIPADA: DECIMAL
                                 */
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                String.format(
                                                        "%.2f",
                                                        minutosSalida),
                                                ReporteUtil.fuenteTablaData(),
                                                fondo));

                                /*
                                 * ACUMULACIÓN DE TOTALES
                                 */
                                totalSegundosPermiso += segundosPermiso;

                                totalMinutosPermiso += minutosPermiso;

                                totalSegundosSalida += segundosSalida;

                                totalMinutosSalida += minutosSalida;

                                contador++;
                            }
                        }

                        /*
                         * FILA DE TOTALES
                         *
                         * Columnas vacías:
                         * 0 N°
                         * 1 horario fecha
                         * 2 horario hora
                         * 3 timbre fecha
                         * 4 timbre hora
                         * 5 tipo permiso
                         * 6 desde
                         *
                         * Columna 7: TOTAL
                         * Columnas 8 y 9: permiso
                         * Columnas 10 y 11: salida anticipada
                         */
                        for (int i = 0; i < 7; i++) {

                            PdfPCell celdaVacia = ReporteUtil.crearCelda(
                                    "",
                                    ReporteUtil.fuenteTablaData(),
                                    Color.WHITE);

                            celdaVacia.setBorder(
                                    Rectangle.NO_BORDER);

                            tablaSalidas.addCell(
                                    celdaVacia);
                        }

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        "TOTAL",
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_SECUNDARIO));

                        /*
                         * TOTAL DEL PERMISO: TIEMPO
                         */
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        convertirMinutosATiempo(
                                                totalSegundosPermiso / 60.0),
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_SECUNDARIO));

                        /*
                         * TOTAL DEL PERMISO: DECIMAL
                         */
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        String.format(
                                                "%.2f",
                                                totalMinutosPermiso),
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_SECUNDARIO));

                        /*
                         * TOTAL DE SALIDA ANTICIPADA: TIEMPO
                         */
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        convertirMinutosATiempo(
                                                totalSegundosSalida / 60.0),
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_SECUNDARIO));

                        /*
                         * TOTAL DE SALIDA ANTICIPADA: DECIMAL
                         */
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(
                                        String.format(
                                                "%.2f",
                                                totalMinutosSalida),
                                        ReporteUtil.fuenteEncabezadoTablaData(),
                                        COLOR_SECUNDARIO));

                        document.add(
                                tablaSalidas);

                        document.add(
                                Chunk.NEWLINE);
                    }
                }
            }

            /*
             * 6. CIERRE Y RETORNO
             */
            document.close();

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;

        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar ReporteSalidasAnticipadas.pdf",
                    e);

        } finally {

            if (document != null && document.isOpen()) {
                try {
                    document.close();
                } catch (Exception ignore) {
                    // Sin acción.
                }
            }

            if (writer != null) {
                try {
                    writer.close();
                } catch (Exception ignore) {
                    // Sin acción.
                }
            }

            if (baos != null) {
                try {
                    baos.close();
                } catch (Exception ignore) {
                    // Sin acción.
                }
            }
        }
    }

    // =========================
    // XLSX (nuevo)
    // =========================
    public byte[] generarReporteXLSX(
            ReporteSalidasAnticipadasRequest request) {
        /*
         * 0. CONSTANTES
         */
        final String NOMBRE_HOJA = "Salidas_Anticipadas";
        final int FILA_ENCABEZADO = 5;

        // B1:T5
        final int MERGE_FIL_INI = 0;
        final int MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1;
        final int MERGE_COL_FIN = 19;

        final String[] HEADERS = {
                "ITEM",
                "IDENTIFICACIÓN",
                "CÓDIGO",
                "APELLIDO NOMBRE",
                "CIUDAD",
                "SUCURSAL",
                "RÉGIMEN",
                "DEPARTAMENTO",
                "CARGO",
                "FECHA HORARIO",
                "HORA HORARIO",
                "FECHA TIMBRE",
                "HORA TIMBRE",
                "TIPO PERMISO",
                "DESDE",
                "HASTA",
                "PERMISO TIEMPO",
                "PERMISO DECIMAL",
                "SALIDA ANTICIPADA HH:MM:SS",
                "SALIDA ANTICIPADA MINUTOS"
        };

        final int[] ANCHOS = {
                10, // ITEM
                20, // IDENTIFICACIÓN
                15, // CÓDIGO
                28, // APELLIDO NOMBRE
                20, // CIUDAD
                22, // SUCURSAL
                22, // RÉGIMEN
                25, // DEPARTAMENTO
                25, // CARGO
                20, // FECHA HORARIO
                18, // HORA HORARIO
                20, // FECHA TIMBRE
                18, // HORA TIMBRE
                25, // TIPO PERMISO
                25, // DESDE
                25, // HASTA
                20, // PERMISO TIEMPO
                20, // PERMISO DECIMAL
                28, // SALIDA ANTICIPADA TIEMPO
                28 // SALIDA ANTICIPADA MINUTOS
        };

        final boolean[] FILTROS = {
                false, // ITEM
                true, // IDENTIFICACIÓN
                true, // CÓDIGO
                true, // APELLIDO NOMBRE
                true, // CIUDAD
                true, // SUCURSAL
                true, // RÉGIMEN
                true, // DEPARTAMENTO
                true, // CARGO
                true, // FECHA HORARIO
                true, // HORA HORARIO
                true, // FECHA TIMBRE
                true, // HORA TIMBRE
                true, // TIPO PERMISO
                true, // DESDE
                true, // HASTA
                true, // PERMISO TIEMPO
                true, // PERMISO DECIMAL
                true, // SALIDA ANTICIPADA TIEMPO
                true // SALIDA ANTICIPADA MINUTOS
        };

        try (
                XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);

            // Mantener visible la cabecera.
            hoja.createFreezePane(
                    0,
                    FILA_ENCABEZADO + 1);

            /*
             * 1. LOGO
             */
            byte[] logo = UtilExcel.decodificarImagenBase64(
                    request.getLogoBase64());

            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(
                        libro,
                        hoja,
                        logo);
            }

            /*
             * 2. COMBINACIÓN DE CELDAS PARA TÍTULOS
             */
            for (int fila = MERGE_FIL_INI; fila <= MERGE_FIL_FIN; fila++) {
                UtilExcel.combinarCeldas(
                        hoja,
                        fila,
                        fila,
                        MERGE_COL_INI,
                        MERGE_COL_FIN);
            }

            /*
             * 3. TÍTULOS
             */
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);

            UtilExcel.establecerTexto(
                    hoja,
                    0,
                    1,
                    UtilExcel.aMayusculasSeguras(
                            request.getEmpresa()),
                    estiloTitulo);

            UtilExcel.establecerTexto(
                    hoja,
                    1,
                    1,
                    "LISTA DE SALIDAS ANTICIPADAS",
                    estiloTitulo);

            String periodo = "PERIODO DEL REPORTE: "
                    + safe(request.getFechaInicio())
                    + " AL "
                    + safe(request.getFechaFin());

            UtilExcel.establecerTexto(
                    hoja,
                    2,
                    1,
                    periodo,
                    estiloTitulo);

            /*
             * 4. ENCABEZADOS
             */
            Row filaHeader = UtilExcel.asegurarFila(
                    hoja,
                    FILA_ENCABEZADO);

            for (int columna = 0; columna < HEADERS.length; columna++) {
                UtilExcel.establecerTexto(
                        filaHeader,
                        columna,
                        HEADERS[columna],
                        null);
            }

            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);

            UtilExcel.aplicarEstiloAFila(
                    filaHeader,
                    HEADERS.length,
                    estiloEncabezado);

            UtilExcel.establecerAnchosColumnas(
                    hoja,
                    ANCHOS);

            hoja.getRow(FILA_ENCABEZADO)
                    .setHeightInPoints(30f);

            /*
             * 5. CUERPO DEL REPORTE
             */
            int filaDatosInicial = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicial;
            int item = 1;

            double totalMinutosPermiso = 0d;
            double totalMinutosSalida = 0d;

            if (request.getGrupos() != null) {
                for (GrupoSalidasDTO grupo : request.getGrupos()) {

                    if (grupo.getEmpleados() == null) {
                        continue;
                    }

                    for (EmpleadoSalidaDTO empleado : grupo.getEmpleados()) {

                        if (empleado.getSalidas() == null) {
                            continue;
                        }

                        String apellidoNombre = (safe(empleado.getApellido())
                                + " "
                                + safe(empleado.getNombre())).trim();

                        for (SalidaDTO salida : empleado.getSalidas()) {

                            Row fila = UtilExcel.asegurarFila(
                                    hoja,
                                    filaActual++);

                            int columna = 0;

                            /*
                             * HORARIO Y TIMBRE
                             */
                            String[] horario = splitFechaHora(
                                    salida.getFecha_hora_horario());

                            String[] timbre = splitFechaHora(
                                    salida.getFecha_hora_timbre());

                            String fechaHorario = horario[0].isEmpty()
                                    ? ""
                                    : ReporteUtil.formatearFechaConDia(
                                            horario[0]);

                            String fechaTimbre = timbre[0].isEmpty()
                                    ? ""
                                    : ReporteUtil.formatearFechaConDia(
                                            timbre[0]);

                            /*
                             * PERMISO
                             */
                            double minutosPermiso = segundosAMinutosConDecimales(
                                    salida.getDiferencia_permiso());

                            boolean tienePermiso = !safe(
                                    salida.getTipo_permiso()).isEmpty()
                                    || minutosPermiso > 0;

                            String tiempoPermiso = tienePermiso
                                    ? convertirMinutosATiempo(
                                            minutosPermiso)
                                    : "";

                            double decimalPermiso = Math.round(
                                    minutosPermiso * 100.0) / 100.0;

                            /*
                             * SALIDA ANTICIPADA RESTANTE
                             */
                            double minutosSalida = segundosAMinutosConDecimales(
                                    salida.getDiferencia());

                            String tiempoSalida = convertirMinutosATiempo(
                                    minutosSalida);

                            double decimalSalida = Math.round(
                                    minutosSalida * 100.0) / 100.0;

                            /*
                             * DATOS GENERALES
                             */
                            UtilExcel.establecerValor(
                                    fila,
                                    columna++,
                                    item++,
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getIdentificacion()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getCodigo()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    apellidoNombre,
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getCiudad()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getSucursal()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getRegimen()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getDepartamento()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(empleado.getCargo()),
                                    null);

                            /*
                             * HORARIO
                             */
                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    fechaHorario,
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    horario[1],
                                    null);

                            /*
                             * TIMBRE
                             */
                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    fechaTimbre,
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    timbre[1],
                                    null);

                            /*
                             * INFORMACIÓN DEL PERMISO
                             */
                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    safe(salida.getTipo_permiso()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    formatearFechaHoraReporte(
                                            salida.getDesde()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    formatearFechaHoraReporte(
                                            salida.getHasta()),
                                    null);

                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    tiempoPermiso,
                                    null);

                            if (tienePermiso) {
                                UtilExcel.establecerValor(
                                        fila,
                                        columna++,
                                        decimalPermiso,
                                        null);
                            } else {
                                UtilExcel.establecerTexto(
                                        fila,
                                        columna++,
                                        "",
                                        null);
                            }

                            /*
                             * SALIDA ANTICIPADA RESTANTE
                             */
                            UtilExcel.establecerTexto(
                                    fila,
                                    columna++,
                                    tiempoSalida,
                                    null);

                            UtilExcel.establecerValor(
                                    fila,
                                    columna++,
                                    decimalSalida,
                                    null);

                            /*
                             * ACUMULAR TOTALES
                             */
                            totalMinutosPermiso += minutosPermiso;
                            totalMinutosSalida += minutosSalida;
                        }
                    }
                }
            }

            int ultimaFilaDatos = filaActual == filaDatosInicial
                    ? FILA_ENCABEZADO
                    : filaActual - 1;

            /*
             * 6. ESTILOS DEL CUERPO
             */
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(
                    libro);

            CellStyle estiloIzquierdaBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(
                    libro);

            if (ultimaFilaDatos >= filaDatosInicial) {
                // ITEM centrado.
                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicial,
                        ultimaFilaDatos,
                        0,
                        0,
                        estiloCentroBorde,
                        true);

                // Resto de columnas.
                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicial,
                        ultimaFilaDatos,
                        1,
                        HEADERS.length - 1,
                        estiloIzquierdaBorde,
                        true);
            }

            /*
             * 7. TABLA CON FILTROS
             *
             * La fila TOTAL queda fuera de esta tabla para que
             * no se mezcle con los filtros.
             */
            if (ultimaFilaDatos >= filaDatosInicial) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "SalidaAnticipadaReporteTabla",
                        FILA_ENCABEZADO,
                        0,
                        ultimaFilaDatos,
                        HEADERS.length - 1,
                        true,
                        FILTROS);
            }

            /*
             * 8. FILA DE TOTALES GENERALES
             */
            if (ultimaFilaDatos >= filaDatosInicial) {
                int filaTotalIndice = ultimaFilaDatos + 1;

                Row filaTotal = UtilExcel.asegurarFila(
                        hoja,
                        filaTotalIndice);

                UtilExcel.establecerTexto(
                        filaTotal,
                        15,
                        "TOTAL",
                        null);

                UtilExcel.establecerTexto(
                        filaTotal,
                        16,
                        convertirMinutosATiempo(
                                totalMinutosPermiso),
                        null);

                UtilExcel.establecerValor(
                        filaTotal,
                        17,
                        Math.round(
                                totalMinutosPermiso * 100.0) / 100.0,
                        null);

                UtilExcel.establecerTexto(
                        filaTotal,
                        18,
                        convertirMinutosATiempo(
                                totalMinutosSalida),
                        null);

                UtilExcel.establecerValor(
                        filaTotal,
                        19,
                        Math.round(
                                totalMinutosSalida * 100.0) / 100.0,
                        null);

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaTotalIndice,
                        filaTotalIndice,
                        15,
                        19,
                        estiloCentroBorde,
                        true);
            }

            /*
             * 9. GENERACIÓN DEL ARCHIVO
             */
            libro.write(baos);

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;

        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar SalidasAnticipadas.xlsx",
                    e);
        }
    }

    // ===== Helpers locales =====
    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String[] splitFechaHora(String fechaHora) {
        // Devuelve [fecha, hora] siempre
        if (fechaHora == null || !fechaHora.contains(" "))
            return new String[] { "", "" };
        String[] p = fechaHora.split(" ");
        String fecha = p.length > 0 ? p[0] : "";
        String hora = p.length > 1 ? p[1] : "";
        return new String[] { fecha, hora };
    }

    private double segundosAMinutosConDecimales(Double segundos) {
        if (segundos == null)
            return 0d;
        return segundos / 60.0;
    }

    private String convertirMinutosATiempo(Double minutos) {
        if (minutos == null || minutos <= 0)
            return "00:00:00";
        int totalSeg = (int) Math.round(minutos * 60);
        int h = totalSeg / 3600;
        int m = (totalSeg % 3600) / 60;
        int s = totalSeg % 60;
        return String.format("%02d:%02d:%02d", h, m, s);
    }

    private String formatearFechaHoraReporte(
            String fechaHora) {
        String valor = safe(fechaHora);

        if (valor.isEmpty()) {
            return "";
        }

        String[] partes = splitFechaHora(valor);

        if (partes[0].isEmpty()) {
            return valor;
        }

        try {
            java.time.LocalDate fecha = java.time.LocalDate.parse(
                    partes[0]);

            String fechaFormateada = fecha.format(
                    java.time.format.DateTimeFormatter.ofPattern(
                            "dd/MM/yyyy"));

            if (partes[1].isEmpty()) {
                return fechaFormateada;
            }

            return fechaFormateada
                    + " "
                    + partes[1];

        } catch (java.time.format.DateTimeParseException e) {
            return valor;
        }
    }

}
