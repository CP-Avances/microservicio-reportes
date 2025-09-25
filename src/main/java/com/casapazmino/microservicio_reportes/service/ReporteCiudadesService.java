package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Ciudad.CiudadDTO;
import com.casapazmino.microservicio_reportes.model.Ciudad.ReporteCiudadesRequest;
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
public class ReporteCiudadesService {

    // =========================
    //          PDF (SIN CAMBIOS)
    // =========================
    public byte[] generarReportePDF(ReporteCiudadesRequest request) {
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
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE CIUDADES"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(new float[]{2, 4});
            tabla.setSpacingBefore(10f);

            tabla.addCell(ReporteUtil.crearCelda("Provincia", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("Ciudad", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            boolean zebra = false;
            for (CiudadDTO ciudad : request.getCiudades()) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(ciudad.getProvincia(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(ciudad.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
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
    //          XLSX (idéntico al front)
    // =========================
    public byte[] generarReporteXLSX(ReporteCiudadesRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Ciudades");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // 2) Merges B1:E1 ... B5:E5
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 4);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 4);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 4);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 4);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 4);

            // 3) Títulos en B1 y B2 (upper)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE CIUDADES", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "ID", "NOMBRE", "PROVINCIA", "ID_PROVINCIA" };
            int[] anchos          = {    10,   20,      20,        20,            20     };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, nombre, provincia, id_prov])
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            List<CiudadDTO> ciudades = request.getCiudades();
            if (ciudades != null) {
                for (int i = 0; i < ciudades.size(); i++) {
                    CiudadDTO c = ciudades.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null);
                    UtilExcel.establecerValor(r, 1, c.getId(), null);
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(c.getNombre()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(c.getProvincia()), null);
                    UtilExcel.establecerValor(r, 4, c.getId_prov(), null);
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..4 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (A6:En), zebra y AutoFilter
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "CiudadesTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true, true }
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
    //           CSV (dinámico del front → orden fijo aquí)
    // =========================
    public byte[] generarReporteCSV(ReporteCiudadesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados como los keys que usabas: id, nombre, provincia, id_prov
            sb.append("id,nombre,provincia,id_prov\n");

            List<CiudadDTO> items = request.getCiudades();
            if (items != null) {
                for (CiudadDTO c : items) {
                    String id      = c.getId() == null ? "" : String.valueOf(c.getId());
                    String nombre  = c.getNombre() == null ? "" : c.getNombre();
                    String prov    = c.getProvincia() == null ? "" : c.getProvincia();
                    String idProv  = c.getId_prov() == null ? "" : String.valueOf(c.getId_prov());
                    sb.append(csv(id)).append(',')
                      .append(csv(nombre)).append(',')
                      .append(csv(prov)).append(',')
                      .append(csv(idProv)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML (igual a tu xml2js)
    // =========================
    public byte[] generarReporteXML(ReporteCiudadesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Ciudades>\n");

            List<CiudadDTO> items = request.getCiudades();
            if (items != null) {
                for (CiudadDTO c : items) {
                    sb.append("  <ciudad id=\"").append(xml(c.getId())).append("\">\n");
                    sb.append("    <nombre>").append(xml(c.getNombre())).append("</nombre>\n");
                    sb.append("    <provincia>").append(xml(c.getProvincia())).append("</provincia>\n");
                    sb.append("  </ciudad>\n");
                }
            }

            sb.append("</Ciudades>\n");
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
