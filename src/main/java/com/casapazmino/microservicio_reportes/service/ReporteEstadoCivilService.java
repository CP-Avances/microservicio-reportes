package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.EstadoCivil.EstadoCivilDTO;
import com.casapazmino.microservicio_reportes.model.EstadoCivil.ReporteEstadosCivilRequest;
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
import java.util.List;

@Service
public class ReporteEstadoCivilService {

    // =========================
    // PDF (SIN CAMBIOS)
    // =========================
    public byte[] generarReportePDF(ReporteEstadosCivilRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE ESTADOS CIVIL"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(30);
            tabla.setWidths(new float[] { 2, 5 });
            tabla.setSpacingBefore(10f);

            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(
                    ReporteUtil.crearCelda("ESTADO CIVIL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            List<EstadoCivilDTO> lista = request.getEstadosCivil();
            boolean zebra = false;
            for (EstadoCivilDTO e : lista) {
                Color bgColor = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(
                        ReporteUtil.crearCelda(String.valueOf(e.getId()), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getEstadoCivil(), ReporteUtil.fuenteTablaData(), bgColor));
                zebra = !zebra;
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
    // XLSX (idéntico al estilo del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteEstadosCivilRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Estado Civil"); // <- nombre exacto

            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // Merges B1:C1 ... B5:C5
            for (int r = 0; r < 5; r++)
                UtilExcel.combinarCeldas(hoja, r, r, 1, 2);

            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE ESTADOS CIVIL", estiloTitulo);

            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CODIGO", "ESTADO CIVIL" }; // <- labels exactos
            int[] anchos = { 20, 30, 40 }; // <- anchos exactos

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // Ordenar por id ASC como en el front
            List<EstadoCivilDTO> estados = request.getEstadosCivil();
            java.util.List<EstadoCivilDTO> ordenados = new java.util.ArrayList<>(
                    estados == null ? java.util.List.of() : estados);
            ordenados.sort(java.util.Comparator.comparing(e -> e.getId() == null ? Integer.MAX_VALUE : e.getId()));

            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            for (int i = 0; i < ordenados.size(); i++) {
                EstadoCivilDTO e = ordenados.get(i);
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                UtilExcel.establecerValor(r, 1, e.getId(), null); // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(e.getEstadoCivil()), null); // ESTADO CIVIL
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);
                // tabla visible A6:Cn
                UtilExcel.crearTablaEstilizada(
                        hoja, "NivelesTitulosTabla", // nombre como en el front viejo
                        filaEncabezado, 0, ultimaFila, encabezados.length - 1,
                        true, new boolean[] { false, true, true });
            }

            libro.write(baos);
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // CSV
    // =========================
    public byte[] generarReporteCSV(ReporteEstadosCivilRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("id,estado_civil\n"); // <- igual al front viejo

            List<EstadoCivilDTO> items = request.getEstadosCivil();
            if (items != null) {
                for (EstadoCivilDTO e : items) {
                    String id = e.getId() == null ? "" : String.valueOf(e.getId());
                    String desc = e.getEstadoCivil() == null ? "" : e.getEstadoCivil();
                    sb.append(csv(id)).append(',').append(csv(desc)).append('\n');
                }
            }
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    // =========================
    // XML
    // =========================
    public byte[] generarReporteXML(ReporteEstadosCivilRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Estados_Civil>\n"); // <- raíz con guion bajo

            List<EstadoCivilDTO> items = request.getEstadosCivil();
            if (items != null) {
                for (EstadoCivilDTO e : items) {
                    sb.append("  <estado_civil id=\"").append(xml(e.getId())).append("\">\n");
                    sb.append("    <estado_civil>").append(xml(e.getEstadoCivil())).append("</estado_civil>\n"); // <-
                                                                                                                 // etiqueta
                                                                                                                 // repetida
                    sb.append("  </estado_civil>\n");
                }
            }

            sb.append("</Estados_Civil>\n");
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    // =========================
    // Helpers CSV/XML
    // =========================
    private String csv(String v) {
        if (v == null)
            return "";
        boolean quote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return quote ? "\"" + s + "\"" : s;
    }

    private String xml(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }
}
