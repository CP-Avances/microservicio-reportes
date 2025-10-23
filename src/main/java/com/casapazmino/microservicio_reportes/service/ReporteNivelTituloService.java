package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.NivelTitulo.NivelTituloDTO;
import com.casapazmino.microservicio_reportes.model.NivelTitulo.ReporteNivelesTitulosRequest;
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
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@Service
public class ReporteNivelTituloService {

    // METODO QUE GENER EL PDF
    public byte[] generarReporteNivelTituloPDF(ReporteNivelesTitulosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 2f, 6f };

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
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE NIVELES DE TÍTULOS PROFESIONALES"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra     = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(40);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL",  ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            List<NivelTituloDTO> lista = request.getNivelesTitulos();
            boolean zebra = false;
            if (lista != null) {
                for (NivelTituloDTO n : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(n.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(n.getNombre(),             ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla, que el controller lo mapee (posible 400)
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteNivelesTitulos.pdf", e);
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
    // XLSX (calcado al front)
    // =========================
    public byte[] generarReporteNivelTituloXLSX(ReporteNivelesTitulosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Niveles Títulos"); // nombre exacto

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // 2) Merges B1:C1 ... B5:C5
            for (int r = 0; r < 5; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, 1, 2);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE NIVELES DE TÍTULOS PROFESIONALES", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CODIGO", "NOMBRE" };
            int[] anchos = { 20, 30, 40 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Ordenar por id ASC (como OrdenarDatos del front) antes de pintar
            List<NivelTituloDTO> datos = request.getNivelesTitulos();
            List<NivelTituloDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(n -> n.getId() == null ? Long.MAX_VALUE : n.getId()));

            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            for (int i = 0; i < ordenados.size(); i++) {
                NivelTituloDTO n = ordenados.get(i);
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                UtilExcel.establecerValor(r, 1, n.getId(), null); // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(n.getNombre()), null); // NOMBRE
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                // Cuerpo: col 0 centrada; resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);

                // 7) Tabla estilizada (A6:Cn), zebra y filtros (ITEM sin filtro)
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "NivelesTitulosTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true });
            }

            libro.write(baos);
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // CSV (orden simple de keys como en front dinámico)
    // =========================
    public byte[] generarReporteNivelTituloCSV(ReporteNivelesTitulosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("id,nombre\n");

            List<NivelTituloDTO> items = request.getNivelesTitulos();
            if (items != null) {
                for (NivelTituloDTO n : items) {
                    String id = n.getId() == null ? "" : String.valueOf(n.getId());
                    String nom = n.getNombre() == null ? "" : n.getNombre();
                    sb.append(csv(id)).append(',').append(csv(nom)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // XML (igual al front: raíz y nodos)
    // =========================
    public byte[] generarReporteNivelTituloXML(ReporteNivelesTitulosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Niveles_titulos>\n");

            // El front ordenaba antes de exportar XML
            List<NivelTituloDTO> datos = request.getNivelesTitulos();
            List<NivelTituloDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(n -> n.getId() == null ? Long.MAX_VALUE : n.getId()));

            for (NivelTituloDTO n : ordenados) {
                sb.append("  <titulos id=\"").append(xml(n.getId())).append("\">\n");
                sb.append("    <nivel>").append(xml(n.getNombre())).append("</nivel>\n");
                sb.append("  </titulos>\n");
            }

            sb.append("</Niveles_titulos>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
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
