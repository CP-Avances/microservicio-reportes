package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteAtrasos.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellStyle;
import java.io.ByteArrayOutputStream;

import java.awt.Color;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAtrasosService {

    public byte[] generarReportePDF(ReporteAtrasosRequest request) {

        // DRY: anchos y colores usados varias veces
        final float[] WIDTHS_TITULO = { 8, 2 };
        final float[] WIDTHS_ENCAB_11 = {
                0.5f, 1.2f,
                1.2f, 1.2f, 1.2f, 1.3f,
                1.0f, 1.0f, 1.4f, 1.8f,
                2.6f
        };
        final float[] WIDTHS_DATA_13 = {
                0.5f, 1.2f,
                1.2f, 1.2f, 1.2f, 1.3f,
                1.0f, 1.0f, 0.7f, 0.7f, 1.8f,
                1.3f, 1.3f
        };

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
            String titulo = "REPORTE DE ATRASOS - " + request.getOpcionBusqueda();
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Contador global
            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    g -> g.getEmpleados().forEach(e -> contadorGlobal.addAndGet(e.getAtrasos().size())));

            // Tabla título + N° registros
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(WIDTHS_TITULO);
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            // Iteración por grupos / empleados
            for (GrupoAtrasoDTO grupo : request.getGrupos()) {
                for (EmpleadoAtrasoDTO emp : grupo.getEmpleados()) {

                    // Info empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                            emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado
                            .addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable contenedora = new PdfPTable(1);
                    contenedora.setWidthPercentage(100);
                    PdfPCell cont = new PdfPCell(infoEmpleado);
                    cont.setPadding(0);
                    cont.setBorder(Rectangle.BOX);
                    contenedora.addCell(cont);
                    document.add(contenedora);

                    // Encabezado (11 columnas)
                    PdfPTable encabezado = new PdfPTable(11);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(WIDTHS_ENCAB_11);

                    encabezado.addCell(
                            ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("TIPO PERMISO", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("PERMISO", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("TOLERANCIA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("ATRASO", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                    for (int i = 0; i < 2; i++) {
                        encabezado.addCell(
                                ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                        encabezado.addCell(
                                ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorSecundario));
                    }
                    encabezado.setSpacingAfter(0f);
                    document.add(encabezado);

                    // Datos (13 columnas, como en tu código original)
                    PdfPTable tablaData = new PdfPTable(13);
                    tablaData.setWidthPercentage(100);
                    tablaData.setWidths(WIDTHS_DATA_13);

                    int contador = 1;
                    long totalSegs = 0L;
                    double totalMins = 0.0;

                    for (AtrasoDTO atraso : emp.getAtrasos()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                        tablaData.addCell(
                                ReporteUtil.crearCelda(String.valueOf(contador), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(atraso.getFechaHorario()),
                                        ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getHoraHorario(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(atraso.getFechaTimbre()),
                                        ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getHoraTimbre(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getTipo_permiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(atraso.getDesde(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(atraso.getHasta(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(ReporteUtil.crearCelda(atraso.getPermiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(ReporteUtil.crearCelda(atraso.getPermiso(), ReporteUtil.fuenteTexto(), fondo)); // (duplicado
                                                                                                                         // en
                                                                                                                         // tu
                                                                                                                         // diseño)
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getTolerancia(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getTiempoAtraso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getMinutosAtraso(), ReporteUtil.fuenteTexto(), fondo));

                        contador++;
                        totalSegs += convertirTiempoAtrasoASegundos(atraso.getTiempoAtraso());
                        totalMins += convertirMinutos(atraso.getMinutosAtraso());
                    }

                    // Totales al final
                    Color fondoTotal = new Color(230, 240, 255);

                    // Vacías hasta la columna 10 (para que TOTAL quede alineado)
                    for (int i = 0; i < 10; i++) {
                        PdfPCell vacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        vacia.setBorder(Rectangle.NO_BORDER);
                        tablaData.addCell(vacia);
                    }


                    // "TOTAL"
                    tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), fondoTotal));

                    // Tiempo total formateado
                    long h = totalSegs / 3600;
                    long m = (totalSegs % 3600) / 60;
                    long s = totalSegs % 60;
                    String totalFormateado = String.format("%02d:%02d:%02d", h, m, s);

                    // Tiempo y minutos
                    tablaData.addCell(ReporteUtil.crearCelda(totalFormateado, ReporteUtil.fuenteTexto(), fondoTotal));
                    tablaData.addCell(ReporteUtil.crearCelda(
                            String.format("%.2f", totalMins).replace(",", "."),
                            ReporteUtil.fuenteTexto(), fondoTotal));

                    tablaData.setSpacingAfter(10f);
                    document.add(tablaData);
                }
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper lanza IAEx, dejamos que el controller decida si es 400
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteAtrasos.pdf", e);
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

    public byte[] generarReporteAtrasosExcel(ReporteAtrasosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Atrasos";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:P5 (B=1 .. P=15 en 0-based)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 15;

        final String[] HEADERS = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "FECHA HORARIO", "HORA HORARIO",
                "FECHA TIMBRE", "HORA TIMBRE",
                "TOLERANCIA", "ATRASO", "ATRASO MINUTOS"
        };
        final int[] ANCHOS = {
                10, 20, 20, 28,
                20, 20, 20, 20, 20,
                20, 20,
                20, 20,
                20, 20, 20
        };

        // ===========================================================
        // 1) Inicialización de recursos (try-with-resources obligatorio)
        // ===========================================================
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // =====================================
            // 2) Construcción (helpers ya existentes)
            // =====================================

            // 2.1 Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2.2 Merges B1:P5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 2.3 Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);

            String activosInactivos = ("1".equals(safe(request.getOpcionBusqueda())) ||
                    "1".equals(String.valueOf(request.getOpcionBusqueda())))
                            ? "ACTIVOS"
                            : "INACTIVOS";
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE ATRASOS - " + activosInactivos, estiloTitulo);

            String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                    + safe(request.getFechaFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            // 2.4 Encabezados + anchos (fila 6 → idx 5)
            Row fh = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(fh, c, HEADERS[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(fh, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 2.5 Cuerpo (aplanado grupos → empleados → atrasos)
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getGrupos() != null) {
                for (GrupoAtrasoDTO grupo : request.getGrupos()) {
                    if (grupo == null || grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoAtrasoDTO emp : grupo.getEmpleados()) {
                        if (emp == null || emp.getAtrasos() == null)
                            continue;

                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                        String ciudad = firstNonEmpty(safe(emp.getCiudad()), safe(grupo.getCiudad()));
                        String sucursal = firstNonEmpty(safe(emp.getSucursal()), safe(grupo.getSucursal()));

                        for (AtrasoDTO reg : emp.getAtrasos()) {
                            if (reg == null)
                                continue;

                            String fechaHor = safe(reg.getFechaHorario());
                            String horaHor = safe(reg.getHoraHorario());
                            String fechaTim = safe(reg.getFechaTimbre());
                            String horaTim = safe(reg.getHoraTimbre());
                            String toler = safe(reg.getTolerancia()); // "HH:mm:ss" o "00:00:00"
                            String atrasoFmt = safe(reg.getTiempoAtraso()); // "HH:mm:ss"
                            String atrasoMin = normalize2(safe(reg.getMinutosAtraso())); // "xx.yy"

                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int col = 0;

                            UtilExcel.establecerValor(r, col++, item++, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCodigo()), null);
                            UtilExcel.establecerTexto(r, col++, apenom, null);
                            UtilExcel.establecerTexto(r, col++, ciudad, null);
                            UtilExcel.establecerTexto(r, col++, sucursal, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getRegimen()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCargo()), null);

                            UtilExcel.establecerTexto(r, col++, fechaHor, null);
                            UtilExcel.establecerTexto(r, col++, horaHor, null);
                            UtilExcel.establecerTexto(r, col++, fechaTim, null);
                            UtilExcel.establecerTexto(r, col++, horaTim, null);

                            UtilExcel.establecerTexto(r, col++, toler, null);
                            UtilExcel.establecerTexto(r, col++, atrasoFmt, null);
                            UtilExcel.establecerTexto(r, col++, atrasoMin, null);
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 2.6 Estilos de cuerpo (bordes + alineación)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con bordes
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // Texto largo a la izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 3, 3, estiloIzqBorde, true); // APELLIDO
                                                                                                            // NOMBRE
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 7, 9, estiloIzqBorde, true); // DEPTO/CARGO
                // Resto centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 2, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 4, 6, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 10, 15, estiloCentroBorde, true);

                // 2.7 Tabla con filtros (ITEM sin filtro)
                boolean[] filtros = new boolean[HEADERS.length];
                for (int i = 0; i < filtros.length; i++)
                    filtros[i] = true;
                filtros[0] = false; // ITEM sin filtro

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "AtrasosReporteTabla",
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
            throw e; // Validación/entrada inválida → el controller puede mapear a 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Atrasos.xlsx", e); // Interno → 500
        }
    }

    /* ===== Helpers locales ===== */
    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String firstNonEmpty(String a, String b) {
        return (a == null || a.isBlank()) ? (b == null ? "" : b) : a;
    }

    private String normalize2(String v) {
        if (v == null || v.isBlank())
            return "0.00";
        String s = v.replace(",", ".");
        try {
            double d = Double.parseDouble(s);
            return String.format(java.util.Locale.US, "%.2f", d);
        } catch (Exception e) {
            return s;
        }
    }

    public static long convertirTiempoAtrasoASegundos(String tiempo) {
        if (tiempo == null || !tiempo.matches("\\d{2}:\\d{2}:\\d{2}"))
            return 0;
        try {
            String[] partes = tiempo.split(":");
            int h = Integer.parseInt(partes[0]);
            int m = Integer.parseInt(partes[1]);
            int s = Integer.parseInt(partes[2]);
            return h * 3600 + m * 60 + s;
        } catch (Exception e) {
            return 0;
        }
    }

    public static double convertirMinutos(String minutosStr) {
        if (minutosStr == null)
            return 0;
        try {
            return Double.parseDouble(minutosStr.replace(",", "."));
        } catch (Exception e) {
            return 0;
        }
    }

}
