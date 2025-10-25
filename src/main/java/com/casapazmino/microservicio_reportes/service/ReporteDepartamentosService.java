package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Departamento.DepartamentoDTO;
import com.casapazmino.microservicio_reportes.model.Departamento.ReporteDepartamentosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
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
public class ReporteDepartamentosService {

    // =========================
    //          PDF 
    // =========================
    public byte[] generarReportePDF(ReporteDepartamentosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 1.5f, 5f, 4f, 1.5f, 4f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DEPARTAMENTOS"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra     = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(5);
            tabla.setWidthPercentage(90);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO",                       ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("SUCURSAL/ ESTABLECIMIENTO",    ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO",                 ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL",                        ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO SUPERIOR",        ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            List<DepartamentoDTO> lista = request.getDepartamentos();
            boolean zebra = false;
            if (lista != null) {
                for (DepartamentoDTO d : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()),         ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getNomsucursal(),                 ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getNombre(),                      ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getNivel()),       ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getDepartamento_padre(),          ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
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
            throw new ReportBuildException("No se pudo generar ReporteDepartamentos.pdf", e);
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
    //          XLSX (idéntico al ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteDepartamentosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Departamentos";
        final int FILA_ENCABEZADO = 5;

        // Merges EXACTOS (B1:G1 ... B5:G5) => (row 0..4, col 1..6)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 6;

        final String[] HEADERS = {
            "ITEM", "ID_SUCURSALES", "NOMBRE SUCURSAL", "ID", "NOMBRE", "NIVEL", "DEPARTAMENTO SUPERIOR"
        };
        final int[] ANCHOS = { 10, 20, 30, 20, 20, 20, 30 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) Merges EXACTOS (B1:G1 ... B5:G5)
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos en B1 y B2 (upper)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE DEPARTAMENTOS", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id_sucursal, nomsucursal, id, nombre, nivel, departamento_padre])
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;

            List<DepartamentoDTO> deps = request.getDepartamentos();
            if (deps != null) {
                for (int i = 0; i < deps.size(); i++) {
                    DepartamentoDTO d = deps.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                    UtilExcel.establecerValor(r, 0, i + 1, null);                                           // ITEM
                    UtilExcel.establecerValor(r, 1, d.getId_sucursal(), null);                              // ID_SUCURSALES
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(d.getNomsucursal()), null);     // NOMBRE SUCURSAL
                    UtilExcel.establecerValor(r, 3, d.getId(), null);                                       // ID
                    UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(d.getNombre()), null);          // NOMBRE
                    UtilExcel.establecerValor(r, 5, d.getNivel(), null);                                     // NIVEL
                    UtilExcel.establecerValor(r, 6, UtilExcel.nuloComoVacio(d.getDepartamento_padre()), null); // DEP. SUPERIOR
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentroBorde, true);

            // Cuerpo
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 6, estiloIzqBorde, true);   // resto
            }

            // 7) Tabla estilizada (TableStyleMedium16), zebra y AutoFilter (A6:Gn)
            if (ultimaFila >= filaDatosIni) {
                boolean[] filtros = new boolean[] { false, true, true, true, true, true, true }; // ITEM off, resto on
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "DepartamentosTabla",
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
            throw new ReportBuildException("No se pudo generar Departamentos.xlsx", e); // Interno → 500
        }
    }

    
    // =========================
    //           CSV (idéntico al front)
    // =========================
    public byte[] generarReporteCSV(ReporteDepartamentosRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Departamentos.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id_sucursal", "nomsucursal", "id", "nombre", "nivel", "departamento_superior" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<DepartamentoDTO> items = request.getDepartamentos();
            if (items != null && !items.isEmpty()) {
                for (DepartamentoDTO d : items) {
                    String idSuc  = (d.getId_sucursal() == null)        ? "" : String.valueOf(d.getId_sucursal());
                    String nomSuc = (d.getNomsucursal() == null)        ? "" : d.getNomsucursal();
                    String id     = (d.getId() == null)                 ? "" : String.valueOf(d.getId());
                    String nombre = (d.getNombre() == null)             ? "" : d.getNombre();
                    String nivel  = (d.getNivel() == null)              ? "" : String.valueOf(d.getNivel());
                    String depSup = (d.getDepartamento_padre() == null) ? "" : d.getDepartamento_padre();

                    sb.append(UtilCsv.csvEscape(idSuc)).append(DELIM)
                    .append(UtilCsv.csvEscape(nomSuc)).append(DELIM)
                    .append(UtilCsv.csvEscape(id)).append(DELIM)
                    .append(UtilCsv.csvEscape(nombre)).append(DELIM)
                    .append(UtilCsv.csvEscape(nivel)).append(DELIM)
                    .append(UtilCsv.csvEscape(depSup)).append(EOL);
                }
            }

            // Retorno (nunca null)
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Interno → 500
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    
    // =========================
    //            XML (igual al xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteDepartamentosRequest request) {
        final String NOMBRE_REPORTE = "Departamentos.xml";
        final String ROOT_TAG = "Departamentos";
        final String ITEM_TAG = "departamento";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<DepartamentoDTO> items = request.getDepartamentos();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (DepartamentoDTO d : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                    .append(" id=\"").append(UtilXml.xmlEsc(d.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<establecimiento>")
                    .append(UtilXml.xmlEsc(d.getNomsucursal()))
                    .append("</establecimiento>").append(EOL);

                    sb.append(IND).append(IND).append("<departamento>")
                    .append(UtilXml.xmlEsc(d.getNombre()))
                    .append("</departamento>").append(EOL);

                    sb.append(IND).append(IND).append("<nivel>")
                    .append(UtilXml.xmlEsc(d.getNivel()))
                    .append("</nivel>").append(EOL);

                    sb.append(IND).append(IND).append("<departamento_superior>")
                    .append(UtilXml.xmlEsc(d.getDepartamento_padre()))
                    .append("</departamento_superior>").append(EOL);

                    sb.append(IND).append("</").append(ITEM_TAG).append(">").append(EOL);
                }
            }

            sb.append("</").append(ROOT_TAG).append(">").append(EOL);
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

}
