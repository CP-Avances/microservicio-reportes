package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Provincia.ProvinciaDTO;
import com.casapazmino.microservicio_reportes.model.Provincia.ReporteProvinciasRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
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
import java.util.List;

@Service
public class ReporteProvinciasService {

    // =========================
    // PDF
    // =========================
    public byte[] generarReportePDF(ReporteProvinciasRequest request) {

        // ➊ DRY: constantes locales (look & feel NO cambia)
        final String TITULO = "LISTA DE PROVINCIAS";
        final float[] WIDTHS = { 2f, 4f }; // mismas proporciones
        final int WIDTH_PERCENT = 50; // mismo 50%
        final float SPACING_BEFORE = 10f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización de recursos
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4); // mismo tamaño
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (usando helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

            Paragraph titulo = ReporteUtil.crearTituloReporte(TITULO);
            titulo.setAlignment(Element.ALIGN_CENTER);
            document.add(titulo);

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(WIDTH_PERCENT);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(SPACING_BEFORE);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("PAÍS", ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));
            tabla.addCell(
                    ReporteUtil.crearCelda("PROVINCIAS", ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));

            // Cuerpo con zebra
            boolean zebra = false;
            List<ProvinciaDTO> lista = request.getProvincias();
            if (lista != null) {
                for (ProvinciaDTO provincia : lista) {
                    Color fondo = zebra ? COLOR_ZEBRA : null; // null conserva blanco
                    tabla.addCell(ReporteUtil.crearCelda(provincia.getPais(), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(provincia.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → que el controller decida 400 si corresponde
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteProvincias.pdf", e);
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
    // XLSX (igual al ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteProvinciasRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Provincias"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

        // MERGES exactos (B1:E1 ... B5:E5) => (row 0..4, col 1..4)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 4;

        final String[] HEADERS = { "ITEM", "ID", "NOMBRE", "ID_PAIS", "PAIS" };
        final int[] ANCHOS = { 10, 20, 20, 20, 20 };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] { false, true, true, true, true };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:E1 ... B5:E5)
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS en B1 y B2
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1,
                    UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE PROVINCIAS", estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO (datos = [item, id, nombre, id_pais, pais])
            int filaDatosInicio = FILA_ENCABEZADO + 1; // 6 → idx 6
            int filaActual = filaDatosInicio;
            int item = 1;

            List<ProvinciaDTO> provincias = request.getProvincias();
            if (provincias != null) {
                for (ProvinciaDTO p : provincias) {
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, item++, null); // ITEM (secuencial)
                    UtilExcel.establecerValor(r, 1, p.getId(), null); // ID
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(p.getNombre()), null); // NOMBRE
                    UtilExcel.establecerValor(r, 3, p.getId_pais(), null); // ID_PAIS
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(p.getPais()), null); // PAIS
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) ALINEACIONES + BORDES (header centrado; cuerpo col 0 centrado, resto
            // izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..4 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) TABLA estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "ProvinciasTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        FILTROS);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Errores de validación → que el controller mapee a 400
            throw e;
        } catch (Exception e) {
            // Errores internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar Provincias.xlsx", e);
        }
    }

    // =========================
    // CSV (como tu ExportToCSV dinámico)
    // =========================
    public byte[] generarReporteCSV(ReporteProvinciasRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Provincias.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id", "nombre", "id_pais", "pais" };

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
            List<ProvinciaDTO> items = request.getProvincias();
            if (items != null && !items.isEmpty()) {
                for (ProvinciaDTO p : items) {
                    String id = (p == null || p.getId() == null) ? "" : String.valueOf(p.getId());
                    String nombre = (p == null || p.getNombre() == null) ? "" : p.getNombre();
                    String idPais = (p == null || p.getId_pais() == null) ? "" : String.valueOf(p.getId_pais());
                    String pais = (p == null || p.getPais() == null) ? "" : p.getPais();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                            .append(UtilCsv.csvEscape(nombre)).append(DELIM)
                            .append(UtilCsv.csvEscape(idPais)).append(DELIM)
                            .append(UtilCsv.csvEscape(pais)).append(EOL);
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
    // XML (igual a xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteProvinciasRequest request) {
        final String NOMBRE_REPORTE = "Provincias.xml";
        final String ROOT_TAG = "Provincias";
        final String ITEM_TAG = "provincia";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<ProvinciaDTO> items = request.getProvincias();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (ProvinciaDTO p : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(p.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<nombre>")
                            .append(UtilXml.xmlEsc(p.getNombre()))
                            .append("</nombre>").append(EOL);

                    sb.append(IND).append(IND).append("<pais>")
                            .append(UtilXml.xmlEsc(p.getPais()))
                            .append("</pais>").append(EOL);

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
