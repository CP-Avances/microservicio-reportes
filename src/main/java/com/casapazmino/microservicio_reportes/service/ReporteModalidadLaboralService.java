package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ModalidadLaboralDTO;
import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ReporteModalidadLaboralRequest;
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
public class ReporteModalidadLaboralService {

    // =========================
    //          PDF (sin cambios)
    // =========================
    public byte[] generarReportePDF(ReporteModalidadLaboralRequest request) {
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
            document.add(ReporteUtil.crearTituloReporte("MODALIDAD LABORAL"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(new float[]{1, 5});
            tabla.setSpacingBefore(10f);

            tabla.addCell(ReporteUtil.crearCelda("ITEM", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("MODALIDAD LABORAL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            boolean zebra = false;
            List<ModalidadLaboralDTO> lista = request.getModalidades();
            if (lista != null) {
                for (ModalidadLaboralDTO modalidad : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(modalidad.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(modalidad.getDescripcion(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
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
    //          XLSX (idéntico al ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteModalidadLaboralRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Modalidad Laboral");

            // 1) LOGO estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) MERGES EXACTOS (B1:C1 ... B5:C5)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 2); // B1:C1
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 2); // B2:C2
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 2); // B3:C3
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 2); // B4:C4
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 2); // B5:C5

            // 3) TÍTULOS en B1 y B2 (centrado y negrita 14)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE MODALIDAD LABORAL", estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS (fila 6 → índice 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CÓDIGO", "DESCRIPCION" };
            int[] anchos = { 20, 30, 40 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) CUERPO (igual al front: datos = [n++, obj.id, obj.descripcion])
            int filaDatosInicio = filaEncabezado + 1; // 6 → índice 6
            int filaActual = filaDatosInicio;
            int item = 1;

            List<ModalidadLaboralDTO> modalidades = request.getModalidades();
            if (modalidades != null) {
                for (ModalidadLaboralDTO m : modalidades) {
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, item++, null);                                   // ITEM (secuencial)
                    UtilExcel.establecerValor(r, 1, m.getId(), null);                                 // CÓDIGO (id)
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(m.getDescripcion()), null); // DESCRIPCION
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) ALINEACIONES + BORDES (como tu obtenerAlineacionHorizontalEmpleados)
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
            UtilExcel.crearTablaEstilizada(
                    hoja,
                    "RegimenTabla", // mismo nombre que tu front para mantener paridad
                    filaEncabezado, 0,
                    ultimaFila, encabezados.length - 1,
                    true,
                    new boolean[] { false, true, true } // filtro: ITEM off, CÓDIGO/DESCRIPCION on
            );

            libro.write(baos);
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //           CSV (idéntico a tu ExportToCSV)
    // =========================
    public byte[] generarReporteCSV(ReporteModalidadLaboralRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados EXACTOS del front antiguo:
            sb.append("ITEM,MODALIDAD_LABORAL\n");

            List<ModalidadLaboralDTO> items = request.getModalidades();
            if (items != null) {
                for (ModalidadLaboralDTO m : items) {
                    String item = (m.getId() == null) ? "" : String.valueOf(m.getId()); // ITEM = id (como tu ExcelJS CSV)
                    String desc = (m.getDescripcion() == null) ? "" : m.getDescripcion();
                    sb.append(csvEscape(item)).append(',').append(csvEscape(desc)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML (idéntico a tu xml2js)
    // =========================
    public byte[] generarReporteXML(ReporteModalidadLaboralRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Modalidad_laboral>\n");

            List<ModalidadLaboralDTO> items = request.getModalidades();
            if (items == null || items.isEmpty()) {
                sb.append("  <lista>NO DEFINIDO</lista>\n");
            } else {
                for (ModalidadLaboralDTO m : items) {
                    sb.append("  <roles id=\"").append(xmlEsc(m.getId())).append("\">\n");
                    sb.append("    <modalidad_laboral>").append(xmlEsc(m.getDescripcion())).append("</modalidad_laboral>\n");
                    sb.append("  </roles>\n");
                }
            }
            sb.append("</Modalidad_laboral>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //          Helpers locales CSV/XML (moveremos a util luego)
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
