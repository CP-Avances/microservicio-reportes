package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ModalidadLaboralDTO;
import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ReporteModalidadLaboralRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

@Service
public class ReporteModalidadLaboralService {

    // =========================
    // PDF (sin cambios)
    // =========================
    public byte[] generarReportePDF(ReporteModalidadLaboralRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 1f, 5f };

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
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("MODALIDAD LABORAL"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("ITEM", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("MODALIDAD LABORAL", ReporteUtil.fuenteEncabezadoTablaData(),
                    colorPrincipal));

            // Cuerpo (zebra)
            List<ModalidadLaboralDTO> lista = request.getModalidades();
            boolean zebra = false;
            if (lista != null) {
                for (ModalidadLaboralDTO modalidad : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(modalidad.getId()),
                            ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(
                            ReporteUtil.crearCelda(modalidad.getDescripcion(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla, dejamos que el controller mapee (posible 400)
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteModalidadLaboral.pdf", e);
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
    // XLSX (idéntico al ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteModalidadLaboralRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Modalidad Laboral"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5;

        // Merges B1:C1 ... B5:C5 => (row 0..4, col 1..2)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 2;

        final String TITULO_REPORTE = "LISTA DE MODALIDAD LABORAL";

        final String[] HEADERS = { "ITEM", "CÓDIGO", "DESCRIPCION" };
        final int[] ANCHOS = { 20, 30, 40 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener encabezado visible

            // 1) LOGO estándar A1:B5 (si existe)
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) MERGES EXACTOS (B1:C1 ... B5:C5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS en B1 y B2
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, TITULO_REPORTE, estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO (datos = [n++, id, descripcion])
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<ModalidadLaboralDTO> modalidades = request.getModalidades();
            if (modalidades != null) {
                for (ModalidadLaboralDTO m : modalidades) {
                    if (m == null)
                        continue;
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, item++, null); // ITEM
                    UtilExcel.establecerValor(r, 1, m.getId(), null); // CÓDIGO
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(m.getDescripcion()), null);// DESCRIPCION
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) ALINEACIONES + BORDES
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..2 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);
            }

            // 7) TABLA estilizada + filtros (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                boolean[] filtros = new boolean[] { false, true, true };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "RegimenTabla", // se mantiene el nombre histórico para paridad con el front
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
            throw new ReportBuildException("No se pudo generar ModalidadLaboral.xlsx", e); // Interno → 500
        }
    }

    // =========================
    // CSV (idéntico a tu ExportToCSV)
    // =========================
    public byte[] generarReporteCSV(ReporteModalidadLaboralRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "ModalidadLaboral.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "ITEM", "MODALIDAD_LABORAL" };

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
            List<ModalidadLaboralDTO> items = request.getModalidades();
            if (items != null && !items.isEmpty()) {
                for (ModalidadLaboralDTO m : items) {
                    String item = (m == null || m.getId() == null) ? "" : String.valueOf(m.getId()); // ITEM = id
                    String desc = (m == null || m.getDescripcion() == null) ? "" : m.getDescripcion();

                    sb.append(UtilCsv.csvEscape(item)).append(DELIM)
                            .append(UtilCsv.csvEscape(desc)).append(EOL);
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
    // XML (idéntico a tu xml2js)
    // =========================
    public byte[] generarReporteXML(ReporteModalidadLaboralRequest request) {
        final String NOMBRE_REPORTE = "Modalidad_laboral.xml";
        final String ROOT_TAG = "Modalidad_laboral";
        final String ITEM_TAG = "roles";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<ModalidadLaboralDTO> items = request.getModalidades();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (ModalidadLaboralDTO m : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(m.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<modalidad_laboral>")
                            .append(UtilXml.xmlEsc(m.getDescripcion()))
                            .append("</modalidad_laboral>").append(EOL);

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
