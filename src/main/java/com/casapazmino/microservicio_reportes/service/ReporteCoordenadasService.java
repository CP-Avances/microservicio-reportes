package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Coordenada.CoordenadaDTO;
import com.casapazmino.microservicio_reportes.model.Coordenada.ReporteCoordenadasRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;

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
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            //LOGO DE EMPRESA
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            //TITULO DE EMPRESA
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            
            //TITULO DE REPORTE
            document.add(ReporteUtil.crearTituloReporte("Lista de coordenadas geográficas"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            //TABLA
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(60);
            tabla.setWidths(new float[]{1.8f, 3f, 5.1f, 5.1f});
            tabla.setSpacingBefore(10f);

            //ENCABEZADOS DE LA TABLA
            String[] headers = {"Código", "Descripción", "Latitud", "Longitud"};
            for (String col : headers) {
                tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            }

            //FILAS DE LA TABLA
            boolean zebra = false;
            for (CoordenadaDTO c : request.getCoordenadas()) {
                Color fondo = zebra ? zebraColor : null;
                zebra = !zebra;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(c.getId()), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getDescripcion(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLatitud(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLongitud(), ReporteUtil.fuenteTablaData(), fondo));
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //          XLSX (idéntico al front)
    // =========================
    public byte[] generarReporteXLSX(ReporteCoordenadasRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Coordenadas");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges (B1:E1 ... B5:E5) → 5 columnas (A..E)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 4);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 4);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 4);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 4);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 4);

            // 3) Títulos (B1 EMPRESA, B2 LISTA DE COORDENADAS)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE COORDENADAS", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CÓDIGO", "LATITUD", "LONGITUD", "DESCRIPCIÓN" };
            int[] anchos          = {    10,      20,        30,        30,          30     };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, latitud, longitud, descripcion])
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            List<CoordenadaDTO> items = request.getCoordenadas();
            if (items != null) {
                for (int i = 0; i < items.size(); i++) {
                    CoordenadaDTO c = items.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(r, 1, c.getId(), null);
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(c.getLatitud()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(c.getLongitud()), null);
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(c.getDescripcion()), null);
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (TableStyleMedium16), zebra y AutoFilter (A6:En)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "CoordenadasTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true, true } // filtro: ITEM off; resto on
                );
            }

            libro.write(baos);
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
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
