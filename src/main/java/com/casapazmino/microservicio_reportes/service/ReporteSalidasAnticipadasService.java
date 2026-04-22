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

    public byte[] generarReportePDF(ReporteSalidasAnticipadasRequest request) {

        // ➊ DRY: constantes locales (no cambia look & feel)
        final float[] WIDTHS_TITULO_TABLA = { 8f, 2f };
        final float[] WIDTHS_INFO_EMPLEADO = { 4f, 4f, 4f };
        final float[] WIDTHS_SALIDAS = {
                1f, // N°
                1.8f, // Horario fecha
                2f, // Horario hora
                2f, // Timbre fecha
                2f, // Timbre hora
                2.5f, // Tipo permiso
                1.5f, // Desde
                1.5f, // Hasta
                1.4f, // Permiso tiempo
                1.2f, // Permiso decimal
                1.8f, // Salida anticipada tiempo
                1.4f // Salida anticipada decimal
        };
        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_AFTER_TITULO_TABLA = 10f;
        final float SPACING_BEFORE_SALIDAS = 5f;
        final float PADDING_ENCABEZADOS = 5f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();

        final String TITULO = "SALIDAS ANTICIPADAS - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
        final String PERIODO = "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin();

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4, 40, 40, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO));
            document.add(ReporteUtil.crearTituloPeriodo(PERIODO));

            // Total de registros (sumatoria de salidas)
            AtomicInteger contadorGlobal = new AtomicInteger();
            if (request.getGrupos() != null) {
                request.getGrupos().forEach(g -> {
                    if (g.getEmpleados() != null)
                        g.getEmpleados().forEach(
                                e -> contadorGlobal.addAndGet(e.getSalidas() != null ? e.getSalidas().size() : 0));
                });
            }

            // Título de tabla + contador
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(WIDTH_PERCENT_100);
            tituloTabla.setWidths(WIDTHS_TITULO_TABLA);
            tituloTabla.setSpacingAfter(SPACING_AFTER_TITULO_TABLA);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(COLOR_SECUNDARIO);
            celdaTitulo.setPadding(PADDING_ENCABEZADOS);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(COLOR_SECUNDARIO);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(PADDING_ENCABEZADOS);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            int contador = 1;

            if (request.getGrupos() != null) {
                for (GrupoSalidasDTO grupo : request.getGrupos()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoSalidaDTO emp : grupo.getEmpleados()) {

                        // Encabezado info empleado (con borde exterior)
                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                        infoEmpleado.setWidths(WIDTHS_INFO_EMPLEADO);

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                emp.getApellido() + " " + emp.getNombre(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), COLOR_ZEBRA));
                        infoEmpleado
                                .addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), COLOR_ZEBRA));
                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), COLOR_ZEBRA));

                        PdfPTable tablaContenedora = new PdfPTable(1);
                        tablaContenedora.setWidthPercentage(WIDTH_PERCENT_100);
                        PdfPCell contenedor = new PdfPCell(infoEmpleado);
                        contenedor.setPadding(0);
                        contenedor.setBorder(Rectangle.BOX);
                        tablaContenedora.addCell(contenedor);
                        document.add(tablaContenedora);

                        // Tabla de salidas
                        PdfPTable tablaSalidas = new PdfPTable(12);
                        tablaSalidas.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaSalidas.setSpacingBefore(SPACING_BEFORE_SALIDAS);
                        tablaSalidas.setWidths(WIDTHS_SALIDAS);

                        // Encabezado fila 1
                        PdfPCell h1 = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        h1.setRowspan(2);
                        tablaSalidas.addCell(h1);
                        PdfPCell h2 = ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        h2.setColspan(2);
                        tablaSalidas.addCell(h2);
                        PdfPCell h3 = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        h3.setColspan(2);
                        tablaSalidas.addCell(h3);
                        PdfPCell h4 = ReporteUtil.crearCelda("TIPO PERMISO", ReporteUtil.fuenteEncabezado(),
                                COLOR_PRIMARIO);
                        h4.setRowspan(2);
                        tablaSalidas.addCell(h4);
                        PdfPCell h5 = ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        h5.setRowspan(2);
                        tablaSalidas.addCell(h5);
                        PdfPCell h6 = ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        h6.setRowspan(2);
                        tablaSalidas.addCell(h6);
                        PdfPCell h7 = ReporteUtil.crearCelda("PERMISO", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        h7.setColspan(2);
                        tablaSalidas.addCell(h7);
                        PdfPCell h8 = ReporteUtil.crearCelda("SALIDA ANTICIPADA", ReporteUtil.fuenteEncabezado(),
                                COLOR_PRIMARIO);
                        h8.setColspan(2);
                        tablaSalidas.addCell(h8);

                        // Encabezado fila 2
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        // PERMISO
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("TIEMPO", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("DECIMAL", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));

                        // SALIDA ANTICIPADA
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("TIEMPO", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda("DECIMAL", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));

                        long totalSegundos = 0;
                        double totalMinutos = 0d;

                        if (emp.getSalidas() != null) {
                            for (SalidaDTO s : emp.getSalidas()) {
                                Color fondo = (contador % 2 == 0) ? COLOR_ZEBRA : Color.WHITE;

                                String[] horaHorario = s.getFecha_hora_horario() != null
                                        ? s.getFecha_hora_horario().split(" ")
                                        : new String[] { "" };
                                String[] horaTimbre = s.getFecha_hora_timbre() != null
                                        ? s.getFecha_hora_timbre().split(" ")
                                        : new String[] { "" };

                                long segundos = s.getDiferencia() != null ? Math.round(s.getDiferencia()) : 0;
                                long horas = segundos / 3600;
                                long minutos = (segundos % 3600) / 60;
                                long resto = segundos % 60;
                                String tiempoFormateado = String.format("%02d:%02d:%02d", horas, minutos, resto);

                                tablaSalidas.addCell(ReporteUtil.crearCelda(String.valueOf(contador),
                                        ReporteUtil.fuenteTexto(), fondo));
                                tablaSalidas.addCell(ReporteUtil.crearCelda(
                                        ReporteUtil.formatearFechaConDia(horaHorario[0]), ReporteUtil.fuenteTexto(),
                                        fondo));
                                tablaSalidas.addCell(ReporteUtil.crearCelda(
                                        horaHorario.length > 1 ? horaHorario[1] : "", ReporteUtil.fuenteTexto(),
                                        fondo));
                                tablaSalidas.addCell(ReporteUtil.crearCelda(
                                        ReporteUtil.formatearFechaConDia(horaTimbre[0]), ReporteUtil.fuenteTexto(),
                                        fondo));
                                tablaSalidas.addCell(ReporteUtil.crearCelda(
                                        horaTimbre.length > 1 ? horaTimbre[1] : "", ReporteUtil.fuenteTexto(), fondo));

                                ///
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(safe(s.getTipo_permiso()), ReporteUtil.fuenteTexto(),
                                                fondo));
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(safe(s.getDesde()), ReporteUtil.fuenteTexto(), fondo));
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(safe(s.getHasta()), ReporteUtil.fuenteTexto(), fondo));

                                // PERMISO
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(safe(s.getPermiso_tiempo()), ReporteUtil.fuenteTexto(),
                                                fondo));
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(safe(s.getPermiso_decimal()), ReporteUtil.fuenteTexto(),
                                                fondo));

                                // SALIDA ANTICIPADA
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(tiempoFormateado, ReporteUtil.fuenteTexto(), fondo));
                                tablaSalidas.addCell(
                                        ReporteUtil.crearCelda(
                                                String.format("%.2f",
                                                        (s.getDiferencia() != null ? s.getDiferencia() : 0d) / 60.0),
                                                ReporteUtil.fuenteTexto(), fondo));
                                //
                                contador++;
                                totalSegundos += segundos;
                                totalMinutos += (s.getDiferencia() != null ? s.getDiferencia() : 0d) / 60.0;
                            }
                        }

                        // Fila de totales (alineada con 12 columnas)
                        // 9 celdas vacías
                        for (int i = 0; i < 9; i++) {
                            PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                            celdaVacia.setBorder(Rectangle.NO_BORDER);
                            tablaSalidas.addCell(celdaVacia);
                        }
                        tablaSalidas.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE));

                        // Tiempo total HH:mm:ss
                        long horasT = totalSegundos / 3600;
                        long minutosT = (totalSegundos % 3600) / 60;
                        long segRest = totalSegundos % 60;
                        String tiempoTotal = String.format("%02d:%02d:%02d", horasT, minutosT, segRest);

                        // Total en minutos
                        tablaSalidas
                                .addCell(ReporteUtil.crearCelda(tiempoTotal, ReporteUtil.fuenteTexto(), Color.WHITE));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(String.format("%.2f", totalMinutos),
                                ReporteUtil.fuenteTexto(), Color.WHITE));

                        document.add(tablaSalidas);
                        document.add(Chunk.NEWLINE);
                    }
                }
            }

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → el controller decidirá 400 si aplica
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteSalidasAnticipadas.pdf", e);
        } finally {
            // 4) Ciclo de recursos garantizado
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

    // =========================
    // XLSX (nuevo)
    // =========================
    public byte[] generarReporteXLSX(ReporteSalidasAnticipadasRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Salidas_Anticipadas"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

        // MERGES exactos (B1:T5) → (row 0..4, col 1..19)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 19;

        final String[] HEADERS = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "FECHA HORARIO", "HORA HORARIO",
                "FECHA TIMBRE", "HORA TIMBRE",
                "TIPO PERMISO", "DESDE", "HASTA",
                "PERMISO TIEMPO", "PERMISO DECIMAL",
                "SALIDA ANTICIPADA HH:MM:SS", "SALIDA ANTICIPADA MINUTOS"
        };
        final int[] ANCHOS = {
                10, 20, 20, 20,
                20, 20, 20, 20, 20,
                20, 20,
                20, 20,
                25, 15, 15,
                18, 18,
                25, 25
        };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] {
                false, true, true, true,
                true, true, true, true, true,
                true, true,
                true, true,
                true, true
        };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:O5)
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS en B1..B3
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE SALIDAS ANTICIPADAS", estiloTitulo); // B2
            String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                    + safe(request.getFechaFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo); // B3

            // 4) ENCABEZADOS + ANCHOS (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO (aplanado grupos → empleados → salidas)
            int filaDatosIni = FILA_ENCABEZADO + 1; // 6 → idx 6
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getGrupos() != null) {
                for (GrupoSalidasDTO grupo : request.getGrupos()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoSalidaDTO usu : grupo.getEmpleados()) {
                        String apenom = (safe(usu.getApellido()) + " " + safe(usu.getNombre())).trim();
                        if (usu.getSalidas() == null)
                            continue;

                        for (SalidaDTO sal : usu.getSalidas()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int col = 0;

                            // === Cálculos (manteniendo helpers existentes) ===
                            String[] ph = splitFechaHora(sal.getFecha_hora_horario());
                            String[] pt = splitFechaHora(sal.getFecha_hora_timbre());
                            String horaHorario = ph[1];
                            String horaTimbre = pt[1];

                            double minutos = segundosAMinutosConDecimales(sal.getDiferencia());
                            String tiempo = convertirMinutosATiempo(minutos);

                            // === Escritura ===
                            UtilExcel.establecerValor(r, col++, item++, null); // ITEM
                            UtilExcel.establecerTexto(r, col++, safe(usu.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getCodigo()), null);
                            UtilExcel.establecerTexto(r, col++, apenom, null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getCiudad()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getSucursal()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getRegimen()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, col++, safe(usu.getCargo()), null);

                            UtilExcel.establecerTexto(r, col++, ph[0], null); // FECHA HORARIO
                            UtilExcel.establecerTexto(r, col++, horaHorario, null); // HORA HORARIO
                            UtilExcel.establecerTexto(r, col++, pt[0], null); // FECHA TIMBRE
                            UtilExcel.establecerTexto(r, col++, horaTimbre, null); // HORA TIMBRE

                            UtilExcel.establecerTexto(r, col++, safe(sal.getTipo_permiso()), null);
                            UtilExcel.establecerTexto(r, col++, safe(sal.getDesde()), null);
                            UtilExcel.establecerTexto(r, col++, safe(sal.getHasta()), null);
                            UtilExcel.establecerTexto(r, col++, safe(sal.getPermiso_tiempo()), null);
                            UtilExcel.establecerTexto(r, col++, safe(sal.getPermiso_decimal()), null);

                            UtilExcel.establecerTexto(r, col++, tiempo, null); // SALIDA ANTICIPADA HH:MM:SS
                            UtilExcel.establecerTexto(r, col++, String.format("%.2f", minutos), null); // SALIDA
                                                                                                       // ANTICIPADA
                                                                                                       // MINUTOS
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) ALINEACIONES + BORDES (header centrado; cuerpo col 0 centrado, resto
            // izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                                                                                                               // centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde,
                        true); // resto izquierda
            }

            // 7) TABLA estilizada + AutoFilter
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "SalidaAnticipadaReporteTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        FILTROS);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación → 400 (lo maneja el controller)
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar SalidasAnticipadas.xlsx", e);
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

}
