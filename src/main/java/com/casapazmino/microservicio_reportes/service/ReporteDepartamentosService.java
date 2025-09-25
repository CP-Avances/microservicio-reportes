package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Departamento.DepartamentoDTO;
import com.casapazmino.microservicio_reportes.model.Departamento.ReporteDepartamentosRequest;
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
public class ReporteDepartamentosService {

    // =========================
    //          PDF (SIN CAMBIOS)
    // =========================
    public byte[] generarReportePDF(ReporteDepartamentosRequest request) {
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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DEPARTAMENTOS"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(5);
            tabla.setWidthPercentage(90);
            tabla.setWidths(new float[]{1.5f, 5, 4, 1.5f, 4});
            tabla.setSpacingBefore(10f);

            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("SUCURSAL/ ESTABLECIMIENTO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO SUPERIOR", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            List<DepartamentoDTO> lista = request.getDepartamentos();
            boolean zebra = false;
            if (lista != null) {
                for (DepartamentoDTO d: lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getNomsucursal(), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getNivel()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getDepartamento_padre(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra= !zebra;
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
    public byte[] generarReporteXLSX(ReporteDepartamentosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Departamentos");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges EXACTOS (B1:G1 ... B5:G5)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 6); // B1:G1
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 6); // B2:G2
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 6); // B3:G3
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 6); // B4:G4
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 6); // B5:G5

            // 3) Títulos en B1 y B2 (upper)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE DEPARTAMENTOS", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = {
                    "ITEM", "ID_SUCURSALES", "NOMBRE SUCURSAL", "ID", "NOMBRE", "NIVEL", "DEPARTAMENTO SUPERIOR"
            };
            int[] anchos = { 10, 20, 30, 20, 20, 20, 30 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id_sucursal, nomsucursal, id, nombre, nivel, departamento_padre])
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            List<DepartamentoDTO> deps = request.getDepartamentos();
            if (deps != null) {
                for (int i = 0; i < deps.size(); i++) {
                    DepartamentoDTO d = deps.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null);                            // ITEM
                    UtilExcel.establecerValor(r, 1, d.getId_sucursal(), null);               // ID_SUCURSALES
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(d.getNomsucursal()), null); // NOMBRE SUCURSAL
                    UtilExcel.establecerValor(r, 3, d.getId(), null);                         // ID
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(d.getNombre()), null); // NOMBRE
                    UtilExcel.establecerValor(r, 5, d.getNivel(), null);                      // NIVEL
                    UtilExcel.establecerValor(r, 6, UtilExcel.nuloComoVacio(d.getDepartamento_padre()), null); // DEP. SUPERIOR
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 6, estiloIzqBorde, true);   // resto
            }

            // 7) Tabla estilizada (TableStyleMedium16), zebra y AutoFilter (A6:Gn)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "DepartamentosTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true, true, true, true } // filtro: ITEM off, resto on
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
    public byte[] generarReporteCSV(ReporteDepartamentosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados EXACTOS y en el mismo orden del front:
            sb.append("id_sucursal,nomsucursal,id,nombre,nivel,departamento_superior\n");

            List<DepartamentoDTO> items = request.getDepartamentos();
            if (items != null) {
                for (DepartamentoDTO d : items) {
                    String idSuc = d.getId_sucursal() == null ? "" : String.valueOf(d.getId_sucursal());
                    String nomSuc= d.getNomsucursal() == null ? "" : d.getNomsucursal();
                    String id    = d.getId() == null ? "" : String.valueOf(d.getId());
                    String nombre= d.getNombre() == null ? "" : d.getNombre();
                    String nivel = d.getNivel() == null ? "" : String.valueOf(d.getNivel());
                    String depSup= d.getDepartamento_padre() == null ? "" : d.getDepartamento_padre();

                    sb.append(csv(idSuc)).append(',')
                      .append(csv(nomSuc)).append(',')
                      .append(csv(id)).append(',')
                      .append(csv(nombre)).append(',')
                      .append(csv(nivel)).append(',')
                      .append(csv(depSup)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML (igual al xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteDepartamentosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Departamentos>\n");

            List<DepartamentoDTO> items = request.getDepartamentos();
            if (items != null) {
                for (DepartamentoDTO d : items) {
                    sb.append("  <departamento id=\"").append(xml(d.getId())).append("\">\n");
                    sb.append("    <establecimiento>").append(xml(d.getNomsucursal())).append("</establecimiento>\n");
                    sb.append("    <departamento>").append(xml(d.getNombre())).append("</departamento>\n");
                    sb.append("    <nivel>").append(xml(d.getNivel())).append("</nivel>\n");
                    sb.append("    <departamento_superior>").append(xml(d.getDepartamento_padre())).append("</departamento_superior>\n");
                    sb.append("  </departamento>\n");
                }
            }

            sb.append("</Departamentos>\n");
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
