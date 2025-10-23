package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Parametro.DetalleParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.Document;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.List;

@Service
public class ReporteParametrosService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteParametrosPDF(ReporteParametrosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS_ENCABEZADO = { 6.9f, 2.1f };
        final float[] WIDTHS_DETALLES   = { 2f, 3f, 6f };
        final String[] HEADERS_DETALLES = { "CÓDIGO DETALLE", "DETALLE", "DESCRIPCIÓN" };

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
            document.add(ReporteUtil.crearTituloReporte("PARÁMETROS GENERALES"));

            // Colores
            Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

            // Iteración de parámetros
            if (request.getParametros() != null) {
                for (ParametroDTO parametro : request.getParametros()) {

                    // Encabezado del bloque de parámetro
                    PdfPTable encabezado = new PdfPTable(2);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(WIDTHS_ENCABEZADO);

                    PdfPCell celdaParametro = new PdfPCell(
                        new Phrase("PARÁMETRO: " + parametro.getDescripcion(), ReporteUtil.fuenteEncabezado())
                    );
                    celdaParametro.setBackgroundColor(colorPrincipal);
                    celdaParametro.setPaddingTop(2f);
                    celdaParametro.setPaddingBottom(3f);
                    celdaParametro.setPaddingLeft(5f);
                    celdaParametro.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                    encabezado.addCell(celdaParametro);

                    PdfPCell celdaCodigo = new PdfPCell(
                        new Phrase("CÓDIGO PARAMETRO: " + parametro.getId(), ReporteUtil.fuenteEncabezado())
                    );
                    celdaCodigo.setBackgroundColor(colorPrincipal);
                    celdaCodigo.setPaddingTop(2f);
                    celdaCodigo.setPaddingBottom(3f);
                    celdaCodigo.setPaddingRight(3f);
                    celdaCodigo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                    encabezado.addCell(celdaCodigo);

                    encabezado.setSpacingAfter(3f);
                    encabezado.setSpacingBefore(12.5f);
                    document.add(encabezado);

                    // Detalles (si existen)
                    if (parametro.getDetalles() != null && !parametro.getDetalles().isEmpty()) {
                        PdfPTable tabla = new PdfPTable(3);
                        tabla.setWidthPercentage(100);
                        tabla.setWidths(WIDTHS_DETALLES);

                        // Encabezados de detalles
                        for (String h : HEADERS_DETALLES) {
                            tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                        }

                        // Cuerpo de detalles (diseño original en blanco)
                        for (DetalleParametroDTO d : parametro.getDetalles()) {
                            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()),      ReporteUtil.fuenteTablaData(), Color.WHITE));
                            tabla.addCell(ReporteUtil.crearCelda(d.getDescripcion(),              ReporteUtil.fuenteTablaData(), Color.WHITE));
                            tabla.addCell(ReporteUtil.crearCelda(d.getObservacion(),              ReporteUtil.fuenteTablaData(), Color.WHITE));
                        }

                        document.add(tabla);
                    }
                }
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla → controller puede mapear 400
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteParametros.pdf", e);
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

    // METODO QUE GENERA EL XLSX
    public byte[] generarReporteParametrosXLSX(ReporteParametrosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Parametros");

            // LOGO
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5 fijo

            // 2) MERGES B1:D1 ... B5:D5
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 3); // B1:D1
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 3); // B2:D2
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 3); // B3:D3
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 3); // B4:D4
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 3); // B5:D5

            // 3) TITULOS (B1 empresa, B2 subtítulo) centrados y negrita 14
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE PARÁMETROS GENERALES", estiloTitulo); // B2

            // 4) ENCABEZADOS y ANCHOS (fila 6 → índice 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CÓDIGO PARAMETRO", "PARÁMETRO", "DETALLE", "DESCRIPCIÓN" };
            int[] anchos = { 10, 25, 50, 20, 160 };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) CUERPO (paridad con tu ExcelJS)
            int filaDatosInicio = filaEncabezado + 1; // 6 → índice 6
            int filaActual = filaDatosInicio;
            int item = 1;

            List<ParametroDTO> parametros = request.getParametros();
            if (parametros != null) {
                for (ParametroDTO p : parametros) {
                    boolean tieneDetalles = (p.getDetalles() != null && !p.getDetalles().isEmpty());
                    if (tieneDetalles) {
                        for (DetalleParametroDTO d : p.getDetalles()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                            UtilExcel.establecerValor(r, 0, item++, null); // ITEM
                            UtilExcel.establecerValor(r, 1, p.getId(), null); // CÓDIGO
                            UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(p.getDescripcion()), null);
                            UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(d.getDescripcion()), null);
                            UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(d.getObservacion()), null);
                        }
                    } else {
                        Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                        UtilExcel.establecerValor(r, 0, item++, null);
                        UtilExcel.establecerValor(r, 1, p.getId(), null);
                        UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(p.getDescripcion()), null);
                        UtilExcel.establecerValor(r, 3, "", null);
                        UtilExcel.establecerValor(r, 4, "", null);
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) ALINEACIONES + BORDES
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 y 1 centrado; 2..4 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 1, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 2, 4, estiloIzqBorde, true);
            }

            // 7) TABLA: estilo TableStyleMedium16 + zebra + AutoFilter (filtros: col 1-2
            // off; 3-5 on)
            UtilExcel.crearTablaEstilizada(hoja,
                    "ParametrosTabla",
                    filaEncabezado, 0,
                    ultimaFila, encabezados.length - 1,
                    true,
                    new boolean[] { false, false, true, true, true });

            libro.write(baos);
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public byte[] generarReporteParametrosCSV(ReporteParametrosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados EXACTOS del front antiguo:
            sb.append("n,codigoParametro,parametro,detalle,descripcion\n");

            int n = 1;
            List<ParametroDTO> parametros = request.getParametros();
            if (parametros != null) {
                for (ParametroDTO p : parametros) {
                    List<DetalleParametroDTO> detalles = p.getDetalles();
                    if (detalles != null && !detalles.isEmpty()) {
                        for (DetalleParametroDTO d : detalles) {
                            sb.append(n++).append(',')
                                    .append(csvEscape(String.valueOf(p.getId()))).append(',')
                                    .append(csvEscape(p.getDescripcion())).append(',')
                                    .append(csvEscape(d.getDescripcion())).append(',')
                                    .append(csvEscape(d.getObservacion())).append('\n');
                        }
                    } else {
                        sb.append(n++).append(',')
                                .append(csvEscape(String.valueOf(p.getId()))).append(',')
                                .append(csvEscape(p.getDescripcion())).append(',')
                                .append(csvEscape("")).append(',')
                                .append(csvEscape("")).append('\n');
                    }
                }
            }
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public byte[] generarReporteParametrosXML(ReporteParametrosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<ParametrosGenerales>\n");

            List<ParametroDTO> parametros = request.getParametros();
            if (parametros != null) {
                for (ParametroDTO p : parametros) {
                    sb.append("  <parametro codigo=\"").append(xmlEsc(p.getId())).append("\">\n");
                    sb.append("    <nombre>").append(xmlEsc(p.getDescripcion())).append("</nombre>\n");
                    sb.append("    <detalles>\n");

                    List<DetalleParametroDTO> detalles = p.getDetalles();
                    if (detalles != null) {
                        for (DetalleParametroDTO d : detalles) {
                            sb.append("      <detalle codigo=\"").append(xmlEsc(d.getId())).append("\">\n");
                            // OJO: En tu XML original el nodo hijo también se llama "detalle"
                            sb.append("        <detalle>").append(xmlEsc(d.getDescripcion())).append("</detalle>\n");
                            sb.append("        <descripcion>").append(xmlEsc(d.getObservacion()))
                                    .append("</descripcion>\n");
                            sb.append("      </detalle>\n");
                        }
                    }
                    sb.append("    </detalles>\n");
                    sb.append("  </parametro>\n");
                }
            }
            sb.append("</ParametrosGenerales>\n");

            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private String xmlEsc(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    private String csvEscape(String v) {
        if (v == null)
            return "";
        boolean mustQuote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return mustQuote ? "\"" + s + "\"" : s;
    }

}
