package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.TimbresLibres.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.Collections;
import java.util.List;

@Service
public class ReporteTimbresLibresService {

    // =========================
    // PDF
    // =========================
    public byte[] generarReportePDF(ReporteTimbresLibresRequest request) {

        // ➊ DRY: constantes locales (manteniendo el look & feel)
        final String TITULO_DEF = "TIMBRES LIBRES - " + ((request.getOpcionBusqueda() != null && request.getOpcionBusqueda() == 1) ? "ACTIVOS" : "INACTIVOS");
        final float[] WIDTHS_CABECERA = { 3f, 3f, 2f };
        final float[] WIDTHS_EMPLEADO = { 3f, 4f, 3f };
        final float[] WIDTHS_TABLA_CON_DISP = { 0.8f, 1.2f, 1.0f, 1.2f, 1.0f, 1.0f, 1.2f, 3.0f, 1.2f, 1.2f };
        final float[] WIDTHS_TABLA_SIN_DISP = { 0.8f, 1.2f, 1.0f, 1.0f, 1.2f, 3.0f, 1.2f, 1.2f };
        final int WIDTH_PERCENT_100 = 100;

        final Color COLOR_PRIMARIO   = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Font  FUENTE_TEXTO     = ReporteUtil.fuenteTexto();

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            final boolean conDispositivo = hayColumnaDispositivo(request);
            final Rectangle pageSize = conDispositivo ? PageSize.A4.rotate() : PageSize.A4;

            baos = new ByteArrayOutputStream();
            document = new Document(pageSize, 40, 40, 50, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            // Empresa
            Paragraph empresa = new Paragraph(safe(request.getEmpresa()), ReporteUtil.fuenteEncabezado());
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            String tituloStr = (request.getTitulo() == null || request.getTitulo().isEmpty()) ? TITULO_DEF : request.getTitulo();
            Paragraph titulo = new Paragraph(tituloStr, ReporteUtil.fuenteEncabezado());
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(0f);
            document.add(titulo);

            // Periodo (si aplica)
            if (request.getPeriodo() != null) {
                Paragraph periodo = new Paragraph(
                        "PERIODO DEL: " + safe(request.getPeriodo().getInicio()) + " AL " + safe(request.getPeriodo().getFin()),
                        ReporteUtil.fuenteTexto()
                );
                periodo.setAlignment(Element.ALIGN_CENTER);
                periodo.setSpacingAfter(0f);
                document.add(periodo);
            }

            // Datos
            if (request.getDatos() != null) {
                for (DatoGrupoDTO grupo : request.getDatos()) {
                    // 2.1 Cabecera principal (3 celdas, borde externo)
                    String descripcion   = resolverDescripcionCabecera(request.getTipoFiltro(), grupo);
                    String establecimiento = esEmpleadoFiltro(request.getTipoFiltro()) ? "" : "SUCURSAL: " + safe(grupo.getSucursal());
                    int totalRegistros   = contarRegistrosTimbres(grupo);

                    PdfPTable tablaCabecera = new PdfPTable(3);
                    tablaCabecera.setWidthPercentage(WIDTH_PERCENT_100);
                    tablaCabecera.setWidths(WIDTHS_CABECERA);
                    tablaCabecera.setSpacingBefore(10f);
                    tablaCabecera.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                    tablaCabecera.addCell(celdaSinBordeIzquierda(descripcion,       FUENTE_TEXTO, COLOR_SECUNDARIO));
                    tablaCabecera.addCell(celdaSinBordeIzquierda(establecimiento,   FUENTE_TEXTO, COLOR_SECUNDARIO));
                    tablaCabecera.addCell(celdaSinBordeIzquierda("N° Registros: " + totalRegistros, FUENTE_TEXTO, COLOR_SECUNDARIO));

                    // Borde alrededor de toda la cabecera (sin tocar helpers)
                    tablaCabecera.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                        PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                        cb.rectangle(
                                widths[0][0],
                                heights[heights.length - 1],
                                widths[0][widths[0].length - 1] - widths[0][0],
                                heights[0] - heights[heights.length - 1]
                        );
                        cb.stroke();
                    });
                    document.add(tablaCabecera);

                    // 2.2 Ficha del empleado (3x3, fondo gris claro, borde externo)
                    if (grupo.getEmpleados() == null) continue;

                    for (EmpleadoDTO empl : grupo.getEmpleados()) {
                        PdfPTable tablaEmpleado = new PdfPTable(3);
                        tablaEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaEmpleado.setSpacingBefore(6f);
                        tablaEmpleado.setWidths(WIDTHS_EMPLEADO);

                        String[][] filas = new String[][] {
                            { "C.C.: " + safe(empl.getIdentificacion()),
                            "EMPLEADO: " + (safe(empl.getApellido()) + " " + safe(empl.getNombre())).trim(),
                            "COD: " + safe(empl.getCodigo()) },
                            { "RÉGIMEN LABORAL: " + safe(empl.getRegimen()),
                            "DEPARTAMENTO: " + safe(empl.getDepartamento()),
                            "CARGO: " + safe(empl.getCargo()) },
                            { "CIUDAD: " + safe(empl.getCiudad()),
                            "SUCURSAL: " + safe(empl.getSucursal()),
                            "ROL: " + safe(empl.getRol()) }
                        };

                        for (String[] fila : filas) {
                            for (String texto : fila) {
                                PdfPCell celda = new PdfPCell(new Phrase(texto, FUENTE_TEXTO));
                                celda.setBackgroundColor(Color.LIGHT_GRAY);
                                celda.setHorizontalAlignment(Element.ALIGN_LEFT);
                                celda.setBorder(Rectangle.NO_BORDER);
                                tablaEmpleado.addCell(celda);
                            }
                        }

                        tablaEmpleado.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                            PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                            cb.rectangle(
                                    widths[0][0],
                                    heights[heights.length - 1],
                                    widths[0][widths[0].length - 1] - widths[0][0],
                                    heights[0] - heights[heights.length - 1]
                            );
                            cb.stroke();
                        });
                        document.add(tablaEmpleado);

                        // 2.3 Tabla de timbres
                        final int cols = conDispositivo ? 10 : 8;
                        PdfPTable tabla = new PdfPTable(cols);
                        tabla.setWidthPercentage(WIDTH_PERCENT_100);
                        tabla.setWidths(conDispositivo ? WIDTHS_TABLA_CON_DISP : WIDTHS_TABLA_SIN_DISP);

                        // Encabezado fila 1
                        addHeaderCellRowSpan(tabla, "N°", COLOR_PRIMARIO, 2);
                        addHeaderCellColSpan(tabla, "TIMBRE", COLOR_PRIMARIO, 2);
                        if (conDispositivo) addHeaderCellColSpan(tabla, "DISPOSITIVO", COLOR_PRIMARIO, 2);
                        addHeaderCellRowSpan(tabla, "RELOJ", COLOR_PRIMARIO, 2);
                        addHeaderCellRowSpan(tabla, "ACCIÓN", COLOR_PRIMARIO, 2);
                        addHeaderCellRowSpan(tabla, "OBSERVACIÓN", COLOR_PRIMARIO, 2);
                        addHeaderCellRowSpan(tabla, "LONGITUD", COLOR_PRIMARIO, 2);
                        addHeaderCellRowSpan(tabla, "LATITUD", COLOR_PRIMARIO, 2);

                        // Encabezado fila 2 (subtítulos)
                        addHeaderCell(tabla, "FECHA", COLOR_PRIMARIO);
                        addHeaderCell(tabla, "HORA",  COLOR_PRIMARIO);
                        if (conDispositivo) {
                            addHeaderCell(tabla, "FECHA", COLOR_PRIMARIO);
                            addHeaderCell(tabla, "HORA",  COLOR_PRIMARIO);
                        }

                        // Cuerpo con zebra
                        int c = 0;
                        List<TimbreDTO> timbres = (empl.getTimbres() != null) ? empl.getTimbres() : Collections.emptyList();
                        if (!timbres.isEmpty()) {
                            for (TimbreDTO t : timbres) {
                                c++;
                                Color bg = (c % 2 == 0) ? new Color(0xE5, 0xE7, 0xE9) : Color.WHITE;

                                addBodyCell(tabla, String.valueOf(c), bg, true);
                                addBodyCell(tabla, safe(t.getFechaServidor()), bg, false);
                                addBodyCell(tabla, safe(t.getHoraServidor()),  bg, false);
                                if (conDispositivo) {
                                    addBodyCell(tabla, safe(t.getFechaDispositivo()), bg, false);
                                    addBodyCell(tabla, safe(t.getHoraDispositivo()),  bg, false);
                                }
                                addBodyCell(tabla, safe(t.getId_reloj()),      bg, true);
                                addBodyCell(tabla, mapAccion(safe(t.getAccion())), bg, true);
                                addBodyCell(tabla, safe(t.getObservacion()),   bg, false);
                                addBodyCell(tabla, safe(t.getLongitud()),      bg, false);
                                addBodyCell(tabla, safe(t.getLatitud()),       bg, false);
                            }
                        } else {
                            PdfPCell vacio = new PdfPCell(new Phrase("SIN REGISTROS", ReporteUtil.fuenteTexto()));
                            vacio.setColspan(cols);
                            vacio.setHorizontalAlignment(Element.ALIGN_CENTER);
                            tabla.addCell(vacio);
                        }

                        document.add(tabla);
                    }
                }
            }

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → que el controller decida 400 si corresponde
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteTimbresLibres.pdf", e);
        } finally {
            // 4) Ciclo de recursos garantizado
            if (document != null && document.isOpen()) {
                try { document.close(); } catch (Exception ignore) {}
            }
            if (writer != null) {
                try { writer.close(); } catch (Exception ignore) {}
            }
            if (baos != null) {
                try { baos.close(); } catch (Exception ignore) {}
            }
        }
    }

    
    // =========================
    // XLSX (misma estructura ExcelJS)
    // =========================
    public byte[] generarReporteXLSX(ReporteTimbresLibresRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Timbres";  // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;          // fila 6 (idx 5)

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // ¿Se muestran columnas de fecha/hora del dispositivo?
            final boolean conDispositivo = hayColumnaDispositivo(request);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges de cabecera
            //    Si hay dispositivo → B..R (col 1..17); si no → B..P (col 1..15)
            final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
            final int MERGE_COL_INI = 1;
            final int MERGE_COL_FIN = conDispositivo ? 17 : 15;
            for (int fila = MERGE_FIL_INI; fila <= MERGE_FIL_FIN; fila++) {
                UtilExcel.combinarCeldas(hoja, fila, fila, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1,
                    UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);

            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? "LISTA DE TIMBRES LIBRES"
                    : request.getTitulo();
            UtilExcel.establecerTexto(hoja, 1, 1,
                    UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);

            if (request.getPeriodo() != null) {
                String periodo = "PERIODO DEL REPORTE: "
                        + safe(request.getPeriodo().getInicio())
                        + " AL "
                        + safe(request.getPeriodo().getFin());
                UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);
            }

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final String[] HEADERS_SIN_DISP = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "FECHA TIMBRE", "HORA TIMBRE", "RELOJ", "ACCIÓN",
                "OBSERVACIÓN", "LATITUD", "LONGITUD"
            };
            final int[] ANCHOS_SIN_DISP = {
                10, 20, 20, 28, 18, 18, 18, 20, 18, 20, 16, 16, 20, 24, 18, 18
            };
            final boolean[] FILTROS_SIN_DISP = {
                false, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true
            };

            final String[] HEADERS_CON_DISP = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "FECHA TIMBRE", "HORA TIMBRE", "RELOJ", "ACCIÓN",
                "OBSERVACIÓN", "LATITUD", "LONGITUD",
                "FECHA TIMBRE DISPOSITIVO", "HORA TIMBRE DISPOSITIVO"
            };
            final int[] ANCHOS_CON_DISP = {
                10, 20, 20, 28, 18, 18, 18, 20, 18, 20, 16, 16, 20, 24, 18, 18, 22, 20
            };
            final boolean[] FILTROS_CON_DISP = {
                false, true, true, true, true, true, true, true, true,
                true, true, true, true, true, true, true, true, true
            };

            final String[] HEADERS = conDispositivo ? HEADERS_CON_DISP : HEADERS_SIN_DISP;
            final int[]    ANCHOS  = conDispositivo ? ANCHOS_CON_DISP  : ANCHOS_SIN_DISP;
            final boolean[] FILTROS = conDispositivo ? FILTROS_CON_DISP : FILTROS_SIN_DISP;

            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupo → empleado → timbre)
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getDatos() != null) {
                for (DatoGrupoDTO grupo : request.getDatos()) {
                    if (grupo == null || grupo.getEmpleados() == null) continue;

                    for (EmpleadoDTO e : grupo.getEmpleados()) {
                        String apenom = (safe(e.getApellido()) + " " + safe(e.getNombre())).trim();
                        List<TimbreDTO> timbres = e.getTimbres();

                        if (timbres == null || timbres.isEmpty()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int c = 0;
                            UtilExcel.establecerValor(r, c++, item++, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCodigo()), null);
                            UtilExcel.establecerTexto(r, c++, apenom, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCiudad()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getSucursal()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getRegimen()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCargo()), null);
                            // columnas de timbre vacías
                            UtilExcel.establecerTexto(r, c++, "", null); // FECHA TIMBRE
                            UtilExcel.establecerTexto(r, c++, "", null); // HORA TIMBRE
                            UtilExcel.establecerTexto(r, c++, "", null); // RELOJ
                            UtilExcel.establecerTexto(r, c++, "", null); // ACCIÓN
                            UtilExcel.establecerTexto(r, c++, "", null); // OBSERVACIÓN
                            UtilExcel.establecerTexto(r, c++, "", null); // LAT
                            UtilExcel.establecerTexto(r, c++, "", null); // LON
                            if (conDispositivo) {
                                UtilExcel.establecerTexto(r, c++, "", null); // FECHA TIMBRE DISPOSITIVO
                                UtilExcel.establecerTexto(r, c++, "", null); // HORA TIMBRE DISPOSITIVO
                            }
                            continue;
                        }

                        for (TimbreDTO t : timbres) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int c = 0;
                            UtilExcel.establecerValor(r, c++, item++, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCodigo()), null);
                            UtilExcel.establecerTexto(r, c++, apenom, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCiudad()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getSucursal()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getRegimen()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCargo()), null);

                            UtilExcel.establecerTexto(r, c++, fechaCortaExcel(t.getFechaServidor()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getHoraServidor()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getId_reloj()), null);
                            UtilExcel.establecerTexto(r, c++, mapAccion(safe(t.getAccion())), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getObservacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getLatitud()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getLongitud()), null);
                            if (conDispositivo) {
                                UtilExcel.establecerTexto(r, c++, fechaCortaExcel(t.getFechaDispositivo()), null);
                                UtilExcel.establecerTexto(r, c++, safe(t.getHoraDispositivo()), null);
                            }
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes (por región)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosIni) {
                String tableName = conDispositivo ? "TimbresReporteTabla" : "TimbresAbiertoReporteTabla";
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    tableName,
                    FILA_ENCABEZADO, 0,
                    ultimaFila, HEADERS.length - 1,
                    true,
                    FILTROS
                );
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar TimbresLibres.xlsx", e);
        }
    }

    
    // =========================
    // Helpers (idénticos al otro módulo)
    // =========================
    private PdfPCell celdaSinBordeIzquierda(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setBorder(Rectangle.NO_BORDER);
        celda.setPaddingTop(6f);
        celda.setPaddingBottom(6f);
        return celda;
    }

    private void addHeaderCell(PdfPTable t, String text, Color bg) {
        PdfPCell c = new PdfPCell(new Phrase(text, ReporteUtil.fuenteTexto()));
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setBackgroundColor(bg);
        t.addCell(c);
    }

    private void addHeaderCellRowSpan(PdfPTable t, String text, Color bg, int rowSpan) {
        PdfPCell c = new PdfPCell(new Phrase(text, ReporteUtil.fuenteTexto()));
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setBackgroundColor(bg);
        c.setRowspan(rowSpan);
        t.addCell(c);
    }

    private void addHeaderCellColSpan(PdfPTable t, String text, Color bg, int colSpan) {
        PdfPCell c = new PdfPCell(new Phrase(text, ReporteUtil.fuenteTexto()));
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setBackgroundColor(bg);
        c.setColspan(colSpan);
        t.addCell(c);
    }

    private void addBodyCell(PdfPTable t, String text, Color bg, boolean center) {
        PdfPCell c = new PdfPCell(new Phrase(text, ReporteUtil.fuenteTexto()));
        c.setBackgroundColor(bg);
        c.setHorizontalAlignment(center ? Element.ALIGN_CENTER : Element.ALIGN_LEFT);
        t.addCell(c);
    }


    private boolean hayColumnaDispositivo(ReporteTimbresLibresRequest req) {
        if (req == null || req.getDatos() == null) return false;
        for (DatoGrupoDTO g : req.getDatos()) {
            if (g.getEmpleados() == null) continue;
            for (EmpleadoDTO e : g.getEmpleados()) {
                if (e.getTimbres() == null) continue;
                for (TimbreDTO t : e.getTimbres()) {
                    if ((t.getFechaDispositivo() != null && !t.getFechaDispositivo().trim().isEmpty()) ||
                        (t.getHoraDispositivo()  != null && !t.getHoraDispositivo().trim().isEmpty())) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private String resolverDescripcionCabecera(String tipoFiltro, DatoGrupoDTO g) {
        String tf = safe(tipoFiltro).toLowerCase();
        switch (tf) {
            case "regimen":      return "RÉGIMEN LABORAL: " + safe(g.getDepartamento()); // si tu payload trae 'nombre', cámbialo por g.getNombre()
            case "departamento": return "DEPARTAMENTO: " + safe(g.getDepartamento());
            case "cargo":        return "CARGO: " + safe(g.getDepartamento());           // idem comentario
            case "ciudad":       return "CIUDAD: " + safe(g.getCiudad());
            case "empleado":     return "LISTA EMPLEADOS";
            default:             return "LISTA EMPLEADOS";
        }
    }

    private boolean esEmpleadoFiltro(String tipoFiltro) {
        return safe(tipoFiltro).equalsIgnoreCase("empleado");
    }

    private int contarRegistrosTimbres(DatoGrupoDTO grupo) {
        if (grupo.getEmpleados() == null) return 0;
        int total = 0;
        for (EmpleadoDTO e : grupo.getEmpleados()) {
            total += (e.getTimbres() == null) ? 0 : e.getTimbres().size();
        }
        return total;
    }

    private String mapAccion(String cod) {
        if (cod == null) return "";
        String k = cod.trim().toUpperCase();
        switch (k) {
            case "EOS": return "Entrada o salida";           // EoS normalizado
            case "AES": return "Inicio o fin alimentación";
            case "PES": return "Inicio o fin permiso";
            case "E":   return "Entrada";
            case "S":   return "Salida";
            case "I/A": return "Inicio alimentación";
            case "F/A": return "Fin alimentación";
            case "I/P": return "Inicio permiso";
            case "F/P": return "Fin permiso";
            case "HA":  return "Timbre libre";
            default:    return cod; // si ya viene mapeado desde FE, se respeta
        }
    }

    private String safe(String val) {
        return (val == null || val.equalsIgnoreCase("null")) ? "" : val;
    }


    private String fechaCortaExcel(String iso) {
        if (iso == null) return "";
        String f = iso.trim();
        return (f.length() >= 10) ? f.substring(0, 10) : f;
        }
}
