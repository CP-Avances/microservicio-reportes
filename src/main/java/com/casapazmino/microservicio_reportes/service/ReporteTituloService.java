package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Titulo.ReporteTitulosRequest;
import com.casapazmino.microservicio_reportes.model.Titulo.TituloDTO;
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
public class ReporteTituloService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteTitulosPDF(ReporteTitulosRequest request) {
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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE TÍTULOS PROFESIONALES"));

            // COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // TABLA
            PdfPTable tabla = new PdfPTable(3);
            tabla.setWidthPercentage(60);
            tabla.setWidths(new float[] { 2, 4, 6 });
            tabla.setSpacingBefore(10f);

            // ENCABEZADOS DE LA TABLA
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NOMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // FILAS DE LA TABLA (CUERPO)
            List<TituloDTO> lista = request.getTitulos();
            boolean zebra = false;
            for (TituloDTO t : lista) {
                Color bgColor = zebra ? colorZebra : Color.WHITE;

                tabla.addCell(
                        ReporteUtil.crearCelda(String.valueOf(t.getId()), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(t.getNivel(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(t.getNombre(), ReporteUtil.fuenteTablaData(), bgColor));
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
    // XLSX (idéntico al front)
    // =========================
    public byte[] generarReporteTitulosXLSX(ReporteTitulosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Títulos");

            // 1) Logo A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // 2) Merges B1:D1 ... B5:D5 (3 columnas de texto, A es el logo)
            for (int r = 0; r < 5; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, 1, 3);
            }

            // 3) Títulos (B1 empresa, B2 "Lista de Títulos")
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TÍTULOS", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CÓDIGO", "NIVEL", "TÍTULO" };
            int[] anchos = { 10, 20, 30, 30 };

            Row header = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(header, c, encabezados[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(header, encabezados.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (en el front no se ordenaba explícitamente; respetamos el orden
            // recibido)
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            List<TituloDTO> titulos = request.getTitulos();
            if (titulos != null) {
                for (int i = 0; i < titulos.size(); i++) {
                    TituloDTO t = titulos.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(r, 1, t.getId(), null);
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(t.getNivel()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(t.getNombre()), null);
                }
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
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 3, estiloIzqBorde, true);

                // 7) Tabla estilizada (A6:Dn), zebra y AutoFilter (ITEM sin filtro)
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TitulosTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true, true });
            }

            libro.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // CSV (orden simple de keys)
    // =========================
    public byte[] generarReporteTitulosCSV(ReporteTitulosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados iguales a las keys del front: id, nivel, nombre
            sb.append("id,nivel,nombre\n");

            List<TituloDTO> items = request.getTitulos();
            if (items != null) {
                for (TituloDTO t : items) {
                    String id = t.getId() == null ? "" : String.valueOf(t.getId());
                    String nivel = t.getNivel() == null ? "" : t.getNivel();
                    String nombre = t.getNombre() == null ? "" : t.getNombre();
                    sb.append(csv(id)).append(',')
                            .append(csv(nivel)).append(',')
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
    // XML (igual a xml2js del front)
    // =========================
    public byte[] generarReporteTitulosXML(ReporteTitulosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Titulos>\n");

            List<TituloDTO> items = request.getTitulos();
            if (items != null) {
                for (TituloDTO t : items) {
                    sb.append("  <titulos id=\"").append(xml(t.getId())).append("\">\n");
                    sb.append("    <nivel>").append(xml(t.getNivel())).append("</nivel>\n");
                    sb.append("    <nombre>").append(xml(t.getNombre())).append("</nombre>\n");
                    sb.append("  </titulos>\n");
                }
            }

            sb.append("</Titulos>\n");
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
