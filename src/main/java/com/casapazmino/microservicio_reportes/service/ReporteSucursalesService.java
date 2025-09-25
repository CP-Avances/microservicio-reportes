package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Sucursal.ReporteSucursalesRequest;
import com.casapazmino.microservicio_reportes.model.Sucursal.SucursalDTO;
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
public class ReporteSucursalesService {

    // =========================
    //          PDF (SIN CAMBIOS)
    // =========================
    public byte[] generarReportePDF(ReporteSucursalesRequest request) {
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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE SUCURSALES"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(3);
            tabla.setWidthPercentage(50);
            tabla.setWidths(new float[]{1.5f, 5f, 2f});
            tabla.setSpacingBefore(10f);

            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("SUCURSAL / ESTABLECIMIENTO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("CIUDAD", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            boolean zebra = false;
            List<SucursalDTO> lista = request.getSucursales();
            if (lista != null) {
                for (SucursalDTO sucursal : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(sucursal.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(sucursal.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(sucursal.getDescripcion(), ReporteUtil.fuenteTablaData(), fondo));
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
    public byte[] generarReporteXLSX(ReporteSucursalesRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Sucursales");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges EXACTOS (B1:D1 ... B5:D5)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 3); // B1:D1
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 3); // B2:D2
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 3); // B3:D3
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 3); // B4:D4
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 3); // B5:D5

            // 3) Títulos (B1 EMPRESA, B2 LISTA DE SUCURSALES)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE SUCURSALES", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "ID", "CIUDAD", "NOMBRE" };
            int[] anchos          = {    10,   20,     20,      30  };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, descripcion→CIUDAD, nombre])
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            List<SucursalDTO> sucursales = request.getSucursales();
            if (sucursales != null) {
                for (int i = 0; i < sucursales.size(); i++) {
                    SucursalDTO s = sucursales.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(r, 1, s.getId(), null); // ID
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(s.getDescripcion()), null); // CIUDAD
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(s.getNombre()), null); // NOMBRE
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo: col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 3, estiloIzqBorde, true);   // ID..NOMBRE
            }

            // 7) Tabla estilizada (TableStyleMedium16), zebra y AutoFilter (A6:Dn)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "SucursalesTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true } // filtro: ITEM off; resto on
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
    //           CSV (como el front: id, ciudad, nombre)
    // =========================
    public byte[] generarReporteCSV(ReporteSucursalesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("id,ciudad,nombre\n");

            List<SucursalDTO> items = request.getSucursales();
            if (items != null) {
                for (SucursalDTO s : items) {
                    String id    = s.getId() == null ? "" : String.valueOf(s.getId());
                    String ciudad= s.getDescripcion() == null ? "" : s.getDescripcion(); // descripcion = ciudad
                    String nombre= s.getNombre() == null ? "" : s.getNombre();
                    sb.append(csv(id)).append(',')
                      .append(csv(ciudad)).append(',')
                      .append(csv(nombre)).append('\n');
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
    public byte[] generarReporteXML(ReporteSucursalesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Establecimientos>\n");

            List<SucursalDTO> items = request.getSucursales();
            if (items != null) {
                for (SucursalDTO s : items) {
                    sb.append("  <establecimiento id=\"").append(xml(s.getId())).append("\">\n");
                    sb.append("    <ciudad>").append(xml(s.getDescripcion())).append("</ciudad>\n");
                    // OJO: el front ponía una etiqueta hija llamada también "establecimiento" con el nombre
                    sb.append("    <establecimiento>").append(xml(s.getNombre())).append("</establecimiento>\n");
                    sb.append("  </establecimiento>\n");
                }
            }

            sb.append("</Establecimientos>\n");
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
