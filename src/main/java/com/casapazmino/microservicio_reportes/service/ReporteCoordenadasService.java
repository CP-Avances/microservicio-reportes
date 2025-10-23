package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Coordenada.CoordenadaDTO;
import com.casapazmino.microservicio_reportes.model.Coordenada.ReporteCoordenadasRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
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
public class ReporteCoordenadasService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteCoordenadasRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 1.8f, 3f, 5.1f, 5.1f };
        final String[] HEADERS = { "Código", "Descripción", "Latitud", "Longitud" };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            Document doc = new Document(PageSize.A4); // alias local para claridad, pero mantenemos referencia en 'document'
            document = doc;
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                request.getUsuario(),
                request.getFraseMarcaAgua(),
                request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes)
            // Logo (opcional)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("Lista de coordenadas geográficas"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color zebraColor     = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(60);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            for (String h : HEADERS) {
                tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            }

            // Filas (zebra)
            boolean zebra = false;
            for (CoordenadaDTO c : request.getCoordenadas()) {
                Color fondo = zebra ? zebraColor : Color.WHITE;
                zebra = !zebra;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(c.getId()),    ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getDescripcion(),           ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLatitud(),               ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLongitud(),              ReporteUtil.fuenteTablaData(), fondo));
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla, que el controller decida (posible 400)
            throw e;
        } catch (Exception e) {
            // 500 uniforme
            throw new ReportBuildException("No se pudo generar ReporteCoordenadas.pdf", e);
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
    //          XLSX (idéntico al front)
    // =========================
    public byte[] generarReporteXLSX(ReporteCoordenadasRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Coordenadas";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:E1 ... B5:E5  => (row 0..4, col 1..4)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 4;

        final String[] HEADERS = { "ITEM", "CÓDIGO", "LATITUD", "LONGITUD", "DESCRIPCIÓN" };
        final int[] ANCHOS     = {   10,      20,        30,        30,          30     };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) Merges B1:E1 ... B5:E5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos (B1 EMPRESA, B2 LISTA DE COORDENADAS)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE COORDENADAS", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, latitud, longitud, descripcion])
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;

            List<CoordenadaDTO> items = request.getCoordenadas();
            if (items != null) {
                for (int i = 0; i < items.size(); i++) {
                    CoordenadaDTO c = items.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                    UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(r, 1, c.getId(), null);
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(c.getLatitud()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(c.getLongitud()), null);
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(c.getDescripcion()), null);
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (TableStyleMedium16), zebra y AutoFilter (A6:En)
            if (ultimaFila >= filaDatosIni) {
                boolean[] filtros = new boolean[] { false, true, true, true, true }; // ITEM sin filtro
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "CoordenadasTabla",
                    FILA_ENCABEZADO, 0,
                    ultimaFila, HEADERS.length - 1,
                    true,
                    filtros
                );
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Coordenadas.xlsx", e); // Interno → 500
        }
    }

    // =========================
    //           CSV (orden fijo)
    // =========================
    public byte[] generarReporteCSV(ReporteCoordenadasRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados como en el front: id, latitud, longitud, descripcion
            sb.append("id,latitud,longitud,descripcion\n");

            List<CoordenadaDTO> items = request.getCoordenadas();
            if (items != null) {
                for (CoordenadaDTO c : items) {
                    String id   = c.getId() == null ? "" : String.valueOf(c.getId());
                    String lat  = c.getLatitud() == null ? "" : c.getLatitud();
                    String lon  = c.getLongitud() == null ? "" : c.getLongitud();
                    String desc = c.getDescripcion() == null ? "" : c.getDescripcion();
                    sb.append(csv(id)).append(',')
                      .append(csv(lat)).append(',')
                      .append(csv(lon)).append(',')
                      .append(csv(desc)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML (según front)
    // =========================
    public byte[] generarReporteXML(ReporteCoordenadasRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Coordenadas>\n");

            List<CoordenadaDTO> items = request.getCoordenadas();
            if (items != null) {
                for (CoordenadaDTO c : items) {
                    sb.append("  <codigo id=\"").append(xml(c.getId())).append("\">\n");
                    sb.append("    <descripcion>").append(xml(c.getDescripcion())).append("</descripcion>\n");
                    sb.append("    <latitud>").append(xml(c.getLatitud())).append("</latitud>\n");
                    sb.append("    <longitud>").append(xml(c.getLongitud())).append("</longitud>\n");
                    sb.append("  </codigo>\n");
                }
            }

            sb.append("</Coordenadas>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //       Helpers CSV/XML
    // =========================
    private String csv(String v) {
        if (v == null) return "";
        boolean quote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return quote ? "\"" + s + "\"" : s;
    }

    private String xml(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"","&quot;")
                .replace("'","&apos;");
    }

}
