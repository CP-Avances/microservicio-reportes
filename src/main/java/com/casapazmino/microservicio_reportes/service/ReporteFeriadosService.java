package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Feriado.FeriadoDTO;
import com.casapazmino.microservicio_reportes.model.Feriado.ReporteFeriadosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@Service
public class ReporteFeriadosService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteFeriadosPDF(ReporteFeriadosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 2.3f, 6f, 3.7f, 4f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (helpers existentes)
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE FERIADOS"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(70);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("CÓDIGO", colorPrincipal));
            tabla.addCell(
                    ReporteUtil.celdaEncabezadoTabla("DESCRIPCIÓN", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("FECHA", colorPrincipal));
            tabla.addCell(
                    ReporteUtil.celdaEncabezadoTabla("RECUPERACIÓN", colorPrincipal));

            // Cuerpo (zebra)
            List<FeriadoDTO> lista = request.getFeriados();
            if (lista != null) {
                lista.sort(Comparator.comparing(FeriadoDTO::getId));
                boolean zebra = false;
                for (FeriadoDTO f : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(
                            ReporteUtil.celdaDataCentro(String.valueOf(f.getId()), fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(f.getDescripcion(), fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(f.getFecha(), fondo));
                    tabla.addCell(
                            ReporteUtil.celdaDataCentro(f.getFechaRecuperacion(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // si un helper valida y falla, el controller decide (400)
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteFeriados.pdf", e);
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
    // XLSX (diseño legacy)
    // =========================
    public byte[] generarReporteFeriadosXLSX(ReporteFeriadosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Feriados"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5;

        // Merges exactos (B1:E1 ... B5:E5) => (row 0..4, col 1..4)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 4;

        final String TITULO_REPORTE = "LISTA DE FERIADOS";
        final String[] HEADERS = { "ITEM", "CÓDIGO", "FERIADO", "FECHA", "FECHA_RECUPERA" };
        final int[] ANCHOS = { 10, 20, 20, 20, 30 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener encabezado visible

            // 1) Logo estándar A1:B5 (si existe)
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges B1:E1 ... B5:E5
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos en B1 y B2 (upper)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, TITULO_REPORTE, estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [item, id, descripcion, fecha, fechaRecuperacion])
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<FeriadoDTO> items = request.getFeriados();
            if (items != null) {
                // Orden estable con nulls al final
                items.sort(java.util.Comparator.comparing(
                        FeriadoDTO::getId,
                        java.util.Comparator.nullsLast(java.util.Comparator.naturalOrder())));

                for (FeriadoDTO f : items) {
                    if (f == null)
                        continue;
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, item++, null); // ITEM
                    UtilExcel.establecerValor(r, 1, f.getId(), null); // CÓDIGO
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(f.getDescripcion()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(f.getFecha()), null);
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(f.getFechaRecuperacion()), null);
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (reutilizar estilos)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; resto izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada + filtros (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                boolean[] filtros = new boolean[] { false, true, true, true, true };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "FeriadosTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        filtros);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Feriados.xlsx", e); // Interno → 500
        }
    }

    // =========================
    // CSV
    // =========================
    public byte[] generarReporteFeriadosCSV(ReporteFeriadosRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Feriados.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "codigo", "feriado", "fecha", "fecha_recupera" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0)
                    sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<FeriadoDTO> items = request.getFeriados();
            if (items != null && !items.isEmpty()) {
                items.sort(Comparator.comparing(FeriadoDTO::getId));
                for (FeriadoDTO f : items) {
                    String id = (f.getId() == null) ? "" : String.valueOf(f.getId());
                    String desc = (f.getDescripcion() == null) ? "" : f.getDescripcion();
                    String fecha = (f.getFecha() == null) ? "" : f.getFecha();
                    String recup = (f.getFechaRecuperacion() == null) ? "" : f.getFechaRecuperacion();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                            .append(UtilCsv.csvEscape(desc)).append(DELIM)
                            .append(UtilCsv.csvEscape(fecha)).append(DELIM)
                            .append(UtilCsv.csvEscape(recup)).append(EOL);
                }
            }

            // Retorno (nunca null)
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Interno → 500
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    // =========================
    // XML
    // =========================
    public byte[] generarReporteFeriadosXML(ReporteFeriadosRequest request) {
        final String NOMBRE_REPORTE = "Feriados.xml";
        final String ROOT_TAG = "Feriados";
        final String ITEM_TAG = "feriados";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<FeriadoDTO> items = request.getFeriados();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                List<FeriadoDTO> ordenados = new ArrayList<>(items);
                ordenados.sort(Comparator.comparingLong(f -> f.getId() == null ? Long.MAX_VALUE : f.getId()));

                for (FeriadoDTO f : ordenados) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(f.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<descripcion>")
                            .append(UtilXml.xmlEsc(f.getDescripcion()))
                            .append("</descripcion>").append(EOL);

                    sb.append(IND).append(IND).append("<fecha>")
                            .append(UtilXml.xmlEsc(f.getFecha()))
                            .append("</fecha>").append(EOL);

                    sb.append(IND).append(IND).append("<fec_recuperacion>")
                            .append(UtilXml.xmlEsc(f.getFechaRecuperacion()))
                            .append("</fec_recuperacion>").append(EOL);

                    sb.append(IND).append("</").append(ITEM_TAG).append(">").append(EOL);
                }
            }

            sb.append("</").append(ROOT_TAG).append(">").append(EOL);
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

}
