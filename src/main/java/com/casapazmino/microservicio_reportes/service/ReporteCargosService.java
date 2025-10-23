package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Cargo.CargoDTO;
import com.casapazmino.microservicio_reportes.model.Cargo.ReporteCargosRequest;
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
public class ReporteCargosService {

    // =========================
    //          PDF (SIN CAMBIOS)
    // =========================
    public byte[] generarReportePDF(ReporteCargosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 1f, 4f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4); // respetamos tu tamaño original
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
            document.add(ReporteUtil.crearTituloReporte("LISTA TIPO DE CARGOS"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra     = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);   // respetamos tu diseño
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("ITEM",   ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("CARGOS", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            boolean zebra = false;
            for (CargoDTO cargo : request.getCargos()) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(cargo.getId()), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(cargo.getCargo(),               ReporteUtil.fuenteTablaData(), fondo));
                zebra = !zebra;
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // si algún helper valida y falla, dejamos que el controller lo trate (posible 400)
            throw e;
        } catch (Exception e) {
            // fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar ReporteCargos.pdf", e);
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
    //          XLSX (Como ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteCargosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Tipos de Cargos");

            // 1) LOGO estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) MERGES EXACTOS (B1:C1 ... B5:C5)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 2); // B1:C1
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 2); // B2:C2
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 2); // B3:C3
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 2); // B4:C4
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 2); // B5:C5

            // 3) TÍTULOS en B1 y B2
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TIPOS DE CARGOS", estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS (fila 6 → índice 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CÓDIGO", "CARGO" };
            int[] anchos = { 10, 30, 45 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) CUERPO (datos = [n++, id, cargo])
            int filaDatosInicio = filaEncabezado + 1; // 6 → índice 6
            int filaActual = filaDatosInicio;
            int n = 1;

            List<CargoDTO> cargos = request.getCargos();
            if (cargos != null) {
                for (CargoDTO c : cargos) {
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, n++, null);                                   // ITEM (secuencial)
                    UtilExcel.establecerValor(r, 1, c.getId(), null);                               // CÓDIGO (id)
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(c.getCargo()), null);   // CARGO
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) ALINEACIONES + BORDES (como obtenerAlineacionHorizontalEmpleados del front)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde   = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..2 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);
            }

            // 7) TABLA estilizada (TableStyleMedium16), zebra y AutoFilter (A6:Cn)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TipoCargoTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true } // filtro: ITEM off, CÓDIGO/CARGO on
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
    //           CSV (idéntico al front)
    // =========================
    public byte[] generarReporteCSV(ReporteCargosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados EXACTOS del front antiguo:
            sb.append("ITEM,CARGOS\n");

            List<CargoDTO> items = request.getCargos();
            if (items != null) {
                for (CargoDTO c : items) {
                    String item = (c.getId() == null) ? "" : String.valueOf(c.getId()); // ITEM = id (así estaba en front)
                    String cargo = (c.getCargo() == null) ? "" : c.getCargo();
                    sb.append(csvEscape(item)).append(',').append(csvEscape(cargo)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML (idéntico a xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteCargosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Cargos>\n");

            List<CargoDTO> items = request.getCargos();
            if (items == null || items.isEmpty()) {
                sb.append("  <lista>NO DEFINIDO</lista>\n"); // opcional; lo mantenemos como convención de backend
            } else {
                for (CargoDTO c : items) {
                    sb.append("  <roles id=\"").append(xmlEsc(c.getId())).append("\">\n");
                    sb.append("    <descripcion>").append(xmlEsc(c.getCargo())).append("</descripcion>\n");
                    sb.append("  </roles>\n");
                }
            }

            sb.append("</Cargos>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //       Helpers CSV/XML (locales, hasta mover a util)
    // =========================
    private String csvEscape(String v) {
        if (v == null) return "";
        boolean mustQuote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return mustQuote ? "\"" + s + "\"" : s;
    }

    private String xmlEsc(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"","&quot;")
                .replace("'","&apos;");
    }
}
