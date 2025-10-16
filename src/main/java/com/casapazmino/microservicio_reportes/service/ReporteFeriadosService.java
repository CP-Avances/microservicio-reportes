package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Feriado.FeriadoDTO;
import com.casapazmino.microservicio_reportes.model.Feriado.ReporteFeriadosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
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
import java.nio.charset.StandardCharsets;
import java.util.Comparator;
import java.util.List;

@Service
public class ReporteFeriadosService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteFeriadosPDF(ReporteFeriadosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            // TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // LOGO DE EMPRESA
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // TITULO DE EMPRESA
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

            // TITULO DE REPORTE
            document.add(ReporteUtil.crearTituloReporte("LISTA DE FERIADOS"));

            // COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // TABLA
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(70);
            tabla.setWidths(new float[] { 2.3f, 6, 3.7f, 4 });
            tabla.setSpacingBefore(10f);

            // ENCABEZADOS DE LA TABLA
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(
                    ReporteUtil.crearCelda("DESCRIPCIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(
                    ReporteUtil.crearCelda("RECUPERACIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // FILAS DE LA TABLA (CUERPO)
            List<FeriadoDTO> lista = request.getFeriados();
            lista.sort(Comparator.comparing(FeriadoDTO::getId));
            boolean zebra = false;
            for (FeriadoDTO f : lista) {
                Color bgColor = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(
                        ReporteUtil.crearCelda(String.valueOf(f.getId()), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(f.getDescripcion(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(f.getFecha(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(f.getFechaRecuperacion(), ReporteUtil.fuenteTablaData(), bgColor));
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
    // XLSX (diseño legacy)
    // =========================
    public byte[] generarReporteFeriadosXLSX(ReporteFeriadosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Feriados");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges (B1:E1 ... B5:E5)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 4);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 4);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 4);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 4);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 4);

            // 3) Títulos en B1 y B2 (upper)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE FERIADOS", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CÓDIGO", "FERIADO", "FECHA", "FECHA_RECUPERA" };
            int[] anchos = { 10, 20, 20, 20, 30 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, descripcion, fecha, fechaRecuperacion])
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            List<FeriadoDTO> items = request.getFeriados();
            if (items != null)
                items.sort(Comparator.comparing(FeriadoDTO::getId));

            if (items != null) {
                for (int i = 0; i < items.size(); i++) {
                    FeriadoDTO f = items.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(r, 1, f.getId(), null);
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(f.getDescripcion()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(f.getFecha()), null);
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(f.getFechaRecuperacion()), null);
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; resto izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (A6:En), zebra y AutoFilter
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "FeriadosTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true, true });
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
    public byte[] generarReporteFeriadosCSV(ReporteFeriadosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados (orden como legacy Excel)
            sb.append("codigo,feriado,fecha,fecha_recupera\n");

            List<FeriadoDTO> items = request.getFeriados();
            if (items != null)
                items.sort(Comparator.comparing(FeriadoDTO::getId));

            if (items != null) {
                for (FeriadoDTO f : items) {
                    String id = f.getId() == null ? "" : String.valueOf(f.getId());
                    String desc = f.getDescripcion() == null ? "" : f.getDescripcion();
                    String fecha = f.getFecha() == null ? "" : f.getFecha();
                    String recup = f.getFechaRecuperacion() == null ? "" : f.getFechaRecuperacion();
                    sb.append(csv(id)).append(',')
                            .append(csv(desc)).append(',')
                            .append(csv(fecha)).append(',')
                            .append(csv(recup)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // XML
    // =========================
    public byte[] generarReporteFeriadosXML(ReporteFeriadosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Feriados>\n");

            List<FeriadoDTO> items = request.getFeriados();
            if (items != null)
                items.sort(Comparator.comparing(FeriadoDTO::getId));

            if (items != null) {
                for (FeriadoDTO f : items) {
                    sb.append("  <feriados id=\"").append(xml(f.getId())).append("\">\n");
                    sb.append("    <descripcion>").append(xml(f.getDescripcion())).append("</descripcion>\n");
                    sb.append("    <fecha>").append(xml(f.getFecha())).append("</fecha>\n");
                    sb.append("    <fec_recuperacion>").append(xml(f.getFechaRecuperacion()))
                            .append("</fec_recuperacion>\n");
                    sb.append("  </feriados>\n");
                }
            }

            sb.append("</Feriados>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // Helpers locales
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
