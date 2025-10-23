package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Discapacidad.DiscapacidadDTO;
import com.casapazmino.microservicio_reportes.model.Discapacidad.ReporteDiscapacidadesRequest;
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
public class ReporteDiscapacidadService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteDiscapacidadesPDF(ReporteDiscapacidadesRequest request) {

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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DISCAPACIDADES"));

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
            tabla.addCell(ReporteUtil.crearCelda("NOMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            List<DiscapacidadDTO> lista = request.getDiscapacidades();
            boolean zebra = false;
            if (lista != null) {
                for (DiscapacidadDTO d : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla, que el controller lo maneje (posible 400)
            throw e;
        } catch (Exception e) {
            // Error interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteDiscapacidades.pdf", e);
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
    // XLSX (idéntico al front)
    // =========================
    public byte[] generarReporteDiscapacidadesXLSX(ReporteDiscapacidadesRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Discapacidades");

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
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE DISCAPACIDADES", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CODIGO", "NOMBRE" };
            int[] anchos = { 20, 30, 40 };

            Row header = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(header, c, encabezados[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(header, encabezados.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Ordenar por id (como hacía el front antes de exportar)
            List<DiscapacidadDTO> datos = request.getDiscapacidades();
            List<DiscapacidadDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(d -> d.getId() == null ? Long.MAX_VALUE : d.getId()));

            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            for (int i = 0; i < ordenados.size(); i++) {
                DiscapacidadDTO d = ordenados.get(i);
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                UtilExcel.establecerValor(r, 1, d.getId(), null); // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(d.getNombre()), null); // NOMBRE
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto
            // izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);

                // 7) Tabla estilizada (A6:Cn), zebra y AutoFilter (ITEM sin filtro)
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "DiscapacidadesTabla",
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
    // CSV (keys simples)
    // =========================
    public byte[] generarReporteDiscapacidadesCSV(ReporteDiscapacidadesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("id,nombre\n");

            List<DiscapacidadDTO> items = request.getDiscapacidades();
            if (items != null) {
                for (DiscapacidadDTO d : items) {
                    String id = d.getId() == null ? "" : String.valueOf(d.getId());
                    String nombre = d.getNombre() == null ? "" : d.getNombre();
                    sb.append(csv(id)).append(',')
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
    // XML (igual a xml2js: raíz y nodos)
    // =========================
    public byte[] generarReporteDiscapacidadesXML(ReporteDiscapacidadesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Discapacidades>\n");

            // Ordenar por id antes de exportar
            List<DiscapacidadDTO> datos = request.getDiscapacidades();
            List<DiscapacidadDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(d -> d.getId() == null ? Long.MAX_VALUE : d.getId()));

            for (DiscapacidadDTO d : ordenados) {
                sb.append("  <discapacidad id=\"").append(xml(d.getId())).append("\">\n");
                sb.append("    <nombre>").append(xml(d.getNombre())).append("</nombre>\n");
                sb.append("  </discapacidad>\n");
            }

            sb.append("</Discapacidades>\n");
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
