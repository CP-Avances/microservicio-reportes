package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Provincia.ProvinciaDTO;
import com.casapazmino.microservicio_reportes.model.Provincia.ReporteProvinciasRequest;
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
public class ReporteProvinciasService {

    // =========================
    //          PDF (SIN CAMBIOS)
    // =========================
    public byte[] generarReportePDF(ReporteProvinciasRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE PROVINCIAS"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(new float[]{2, 4});
            tabla.setSpacingBefore(10f);

            tabla.addCell(ReporteUtil.crearCelda("PAÍS", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("PROVINCIAS", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            boolean zebra = false;
            List<ProvinciaDTO> lista = request.getProvincias();
            for (ProvinciaDTO provincia : lista) {
                Color fondo = zebra ? colorZebra : null;
                tabla.addCell(ReporteUtil.crearCelda(provincia.getPais(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(provincia.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                zebra = !zebra;
            }

            document.add(tabla);
            document.close();
            writer.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //          XLSX (igual al ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteProvinciasRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Provincias");

            // 1) LOGO estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) MERGES EXACTOS (B1:E1 ... B5:E5)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 4); // B1:E1
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 4); // B2:E2
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 4); // B3:E3
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 4); // B4:E4
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 4); // B5:E5

            // 3) TÍTULOS en B1 y B2
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE PROVINCIAS", estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS (fila 6 → índice 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "ID", "NOMBRE", "ID_PAIS", "PAIS" };
            int[] anchos =      {     10,   20,      20,        20,      20 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) CUERPO (datos = [index+1, id, nombre, id_pais, pais])
            int filaDatosInicio = filaEncabezado + 1; // 6 → índice 6
            int filaActual = filaDatosInicio;

            List<ProvinciaDTO> provincias = request.getProvincias();
            if (provincias != null) {
                for (int i = 0; i < provincias.size(); i++) {
                    ProvinciaDTO p = provincias.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null);                             // ITEM
                    UtilExcel.establecerValor(r, 1, p.getId(), null);                         // ID
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(p.getNombre()), null); // NOMBRE
                    UtilExcel.establecerValor(r, 3, p.getId_pais(), null);                    // ID_PAIS
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(p.getPais()), null);   // PAIS
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) ALINEACIONES + BORDES (header centrado; cuerpo col 0 centrado, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde   = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..4 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) TABLA estilizada (TableStyleMedium16), zebra y AutoFilter (A6:En)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "ProvinciasTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true, true } // filtro: ITEM off, resto on
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
    //           CSV (como tu ExportToCSV dinámico)
    // =========================
    public byte[] generarReporteCSV(ReporteProvinciasRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados en el mismo orden de las claves principales del objeto:
            sb.append("id,nombre,id_pais,pais\n");

            List<ProvinciaDTO> items = request.getProvincias();
            if (items != null) {
                for (ProvinciaDTO p : items) {
                    String id      = p.getId() == null ? "" : String.valueOf(p.getId());
                    String nombre  = p.getNombre() == null ? "" : p.getNombre();
                    String idPais  = p.getId_pais() == null ? "" : String.valueOf(p.getId_pais());
                    String pais    = p.getPais() == null ? "" : p.getPais();
                    sb.append(csv(id)).append(',')
                      .append(csv(nombre)).append(',')
                      .append(csv(idPais)).append(',')
                      .append(csv(pais)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML (igual a xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteProvinciasRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Provincias>\n");

            List<ProvinciaDTO> items = request.getProvincias();
            if (items != null) {
                for (ProvinciaDTO p : items) {
                    sb.append("  <provincia id=\"").append(xml(p.getId())).append("\">\n");
                    sb.append("    <nombre>").append(xml(p.getNombre())).append("</nombre>\n");
                    sb.append("    <pais>").append(xml(p.getPais())).append("</pais>\n");
                    sb.append("  </provincia>\n");
                }
            }
            sb.append("</Provincias>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //       Helpers CSV/XML (locales por ahora)
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
