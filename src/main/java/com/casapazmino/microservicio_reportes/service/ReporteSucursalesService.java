package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Sucursal.ReporteSucursalesRequest;
import com.casapazmino.microservicio_reportes.model.Sucursal.SucursalDTO;
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
public class ReporteSucursalesService {

    // =========================
    //          PDF
    // =========================
    public byte[] generarReportePDF(ReporteSucursalesRequest request) {

        // ➊ DRY: constantes locales (mismo look & feel)
        final String TITULO = "LISTA DE SUCURSALES";
        final float[] WIDTHS = { 1.5f, 5f, 2f };
        final int WIDTH_PERCENT_50 = 50;
        final float SPACING_BEFORE = 10f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_ZEBRA    = ReporteUtil.colorZebraClaro();

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
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO));

            PdfPTable tabla = new PdfPTable(3);
            tabla.setWidthPercentage(WIDTH_PERCENT_50);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(SPACING_BEFORE);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO",                     ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));
            tabla.addCell(ReporteUtil.crearCelda("SUCURSAL / ESTABLECIMIENTO", ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));
            tabla.addCell(ReporteUtil.crearCelda("CIUDAD",                     ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));

            // Cuerpo con zebra
            boolean zebra = false;
            List<SucursalDTO> lista = request.getSucursales();
            if (lista != null) {
                for (SucursalDTO s : lista) {
                    Color fondo = zebra ? COLOR_ZEBRA : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(s.getId()),  ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(s.getNombre(),              ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(s.getDescripcion(),         ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → el controller podría responder 400
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteSucursales.pdf", e);
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
    //          XLSX
    // =========================
    public byte[] generarReporteXLSX(ReporteSucursalesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Sucursales";  // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;            // fila visual 6 (idx 5)

        // MERGES exactos (B1:D1 ... B5:D5) => (row 0..4, col 1..3)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 3;

        final String[] HEADERS = { "ITEM", "ID", "CIUDAD", "NOMBRE" };
        final int[]    ANCHOS  = {    10,   20,     20,      30  };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] { false, true, true, true };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:D1 ... B5:D5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos (B1 EMPRESA, B2 LISTA DE SUCURSALES)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                hoja, 0, 1,
                UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                estiloTitulo
            ); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE SUCURSALES", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, ciudad=descripcion, nombre])
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<SucursalDTO> sucursales = request.getSucursales();
            if (sucursales != null) {
                for (SucursalDTO s : sucursales) {
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, item++, null);                                    // ITEM
                    UtilExcel.establecerValor(r, 1, s.getId(), null);                                 // ID
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(s.getDescripcion()), null); // CIUDAD
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(s.getNombre()), null);      // NOMBRE
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo: col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header
            UtilExcel.aplicarEstiloARegion(
                hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                estiloCentroBorde, true
            );

            // Cuerpo
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 3, estiloIzqBorde, true);    // ID..NOMBRE
            }

            // 7) Tabla estilizada (zebra) + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "SucursalesTabla",
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
            // Validaciones → 400
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar Sucursales.xlsx", e);
        }
    }

    
    // =========================
    //           CSV (como el front: id, ciudad, nombre)
    // =========================
    public byte[] generarReporteCSV(ReporteSucursalesRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Sucursales.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id", "ciudad", "nombre" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<SucursalDTO> items = request.getSucursales();
            if (items != null && !items.isEmpty()) {
                for (SucursalDTO s : items) {
                    String id     = (s == null || s.getId() == null)           ? "" : String.valueOf(s.getId());
                    String ciudad = (s == null || s.getDescripcion() == null)  ? "" : s.getDescripcion(); // descripcion = ciudad
                    String nombre = (s == null || s.getNombre() == null)       ? "" : s.getNombre();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                    .append(UtilCsv.csvEscape(ciudad)).append(DELIM)
                    .append(UtilCsv.csvEscape(nombre)).append(EOL);
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
    //            XML (igual a xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteSucursalesRequest request) {
        final String NOMBRE_REPORTE = "Sucursales.xml";
        final String ROOT_TAG = "Establecimientos";
        final String ITEM_TAG = "establecimiento";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder();

            // 1) Encabezado
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            // 2) Cuerpo
            List<SucursalDTO> items = (request == null) ? null : request.getSucursales();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (SucursalDTO s : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                    .append(" id=\"").append(UtilXml.xmlEsc(s == null ? null : s.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<ciudad>")
                    .append(UtilXml.xmlEsc(s == null ? null : s.getDescripcion()))
                    .append("</ciudad>").append(EOL);

                    // Importante: mantener etiqueta hija llamada también "establecimiento" (diseño original del front)
                    sb.append(IND).append(IND).append("<establecimiento>")
                    .append(UtilXml.xmlEsc(s == null ? null : s.getNombre()))
                    .append("</establecimiento>").append(EOL);

                    sb.append(IND).append("</").append(ITEM_TAG).append(">").append(EOL);
                }
            }

            // 3) Cierre
            sb.append("</").append(ROOT_TAG).append(">").append(EOL);
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }


}
