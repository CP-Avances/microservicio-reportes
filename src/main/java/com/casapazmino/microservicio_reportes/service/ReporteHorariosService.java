package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Horario.DetalleHorarioDTO;
import com.casapazmino.microservicio_reportes.model.Horario.HorarioDTO;
import com.casapazmino.microservicio_reportes.model.Horario.ReporteHorariosRequest;
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
public class ReporteHorariosService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteHorariosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            //LOGO DE EMPRESA
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            //TITULO DE EMPRESA
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            
            //TITULO DE REPORTE
            document.add(ReporteUtil.crearTituloReporte("LISTA DE HORARIOS"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            for (HorarioDTO h : request.getHorarios()) {
                // Tabla contenedora
                PdfPTable bloque = new PdfPTable(1);
                bloque.setWidthPercentage(100);
                bloque.setSpacingBefore(10f);

                //TABLA CABEZERA DE CADA HORARIO
                PdfPTable cabecera = new PdfPTable(3);
                cabecera.setWidthPercentage(100);
                cabecera.setWidths(new float[]{5, 5, 5});

                cabecera.addCell(celdaHorario("HORARIO: " + h.getNombre(), colorPrincipal, Rectangle.TOP | Rectangle.LEFT));
                cabecera.addCell(celdaHorario("HORAS DE TRABAJO: " + h.getHoraTrabajo(), colorPrincipal, Rectangle.TOP));
                cabecera.addCell(celdaHorario("MINUTOS DE ALIMENTACIÓN: " + h.getMinutosComida(), colorPrincipal, Rectangle.TOP | Rectangle.RIGHT));

                cabecera.addCell(celdaHorario("CÓDIGO: " + h.getCodigo(), colorPrincipal, Rectangle.LEFT));
                cabecera.addCell(celdaHorario("HORARIO NOCTURNO: " + (h.isNoturno() ? "Sí" : "No"), colorPrincipal, Rectangle.NO_BORDER));
                cabecera.addCell(celdaHorario("DOCUMENTO: " + (h.getDocumento() != null ? h.getDocumento() : ""), colorPrincipal, Rectangle.RIGHT));

                PdfPCell celdaCabecera = new PdfPCell(cabecera);
                celdaCabecera.setPadding(0);
                celdaCabecera.setBorder(Rectangle.NO_BORDER);
                bloque.addCell(celdaCabecera);

                if (h.getDetalles() != null && !h.getDetalles().isEmpty()) {
                    // Título "DETALLES"
                    PdfPTable tituloDetalles = new PdfPTable(1);
                    tituloDetalles.setWidthPercentage(100);
                    PdfPCell celdaDetalles = new PdfPCell(new Phrase("DETALLES", ReporteUtil.fuenteEncabezadoTablaData()));
                    celdaDetalles.setBackgroundColor(colorSecundario);
                    celdaDetalles.setHorizontalAlignment(Element.ALIGN_CENTER);
                    celdaDetalles.setPadding(5);
                    celdaDetalles.setBorder(Rectangle.BOX);
                    tituloDetalles.addCell(celdaDetalles);

                    PdfPCell celdaTitulo = new PdfPCell(tituloDetalles);
                    celdaTitulo.setPadding(0);
                    celdaTitulo.setBorder(Rectangle.NO_BORDER);
                    bloque.addCell(celdaTitulo);

                    // Tabla de detalles
                    PdfPTable tabla = new PdfPTable(7);
                    tabla.setWidthPercentage(100);
                    tabla.setWidths(new float[]{1.5f, 1.5f, 2.5f, 4.4f, 2, 3, 3});

                    String[] headers = {"ORDEN", "HORA", "TOLERANCIA", "ACCIÓN", "OTRO DÍA", "MINUTOS ANTES", "MINUTOS DESPUÉS"};
                    for (String col : headers) {
                        tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                    }

                    boolean zebra = false;
                    for (DetalleHorarioDTO d : h.getDetalles()) {
                        Color fondo = zebra ? zebraColor : null;
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getOrden()), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getHora(), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getTolerancia() != null ? d.getTolerancia().toString() : "", ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getTipoAccionShow(), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.isSegundoDia() ? "Sí" : "No", ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getMinutosAntes()), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getMinutosDespues()), ReporteUtil.fuenteTablaData(), fondo));
                        zebra = !zebra;
                    }

                    PdfPCell celdaTabla = new PdfPCell(tabla);
                    celdaTabla.setPadding(0);
                    celdaTabla.setBorder(Rectangle.NO_BORDER);
                    bloque.addCell(celdaTabla);
                }

                document.add(bloque);
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell celdaHorario(String texto, Color fondo, int border) {
        Font fuente = FontFactory.getFont(FontFactory.HELVETICA, 9);
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(5f);
        celda.setBorder(border);
        return celda;
    }

        // =========================
    //           XLSX
    // =========================
    public byte[] generarReporteXLSX(ReporteHorariosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Horarios");

            // 1) Logo A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // 2) Merges B1:N1 ... B5:N5 (N = 14 columnas: A..N)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 13);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 13);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 13);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 13);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 13);

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE HORARIOS", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = {
                    "ITEM","HORARIO","CÓDIGO","HORAS DE TRABAJO","MINUTOS DE ALIMENTACIÓN",
                    "HORARIO NOTURNO","DOCUMENTO","ORDEN","HORA","TOLERANCIA",
                    "ACCIÓN","OTRO DÍA","MINUTOS ANTES","MINUTOS DESPUÉS"
            };
            int[] anchos = {10,20,20,20,20,20,20,20,20,20,20,20,30,30};

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado horarios x detalles)
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<HorarioDTO> horarios = request.getHorarios();
            if (horarios != null) {
                for (HorarioDTO h : horarios) {
                    List<DetalleHorarioDTO> dets = h.getDetalles();
                    if (dets == null || dets.isEmpty()) continue;

                    for (DetalleHorarioDTO d : dets) {
                        Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                        UtilExcel.establecerValor(r, 0, item++, null);                              // ITEM
                        UtilExcel.establecerTexto(r, 1, nvl(h.getNombre()), null);                  // HORARIO
                        UtilExcel.establecerTexto(r, 2, nvl(h.getCodigo()), null);                  // CÓDIGO
                        UtilExcel.establecerTexto(r, 3, nvl(h.getHoraTrabajo()), null);            // HORAS DE TRABAJO
                        UtilExcel.establecerTexto(r, 4, nvl(h.getMinutosComida()), null);          // MINUTOS DE ALIMENTACIÓN
                        UtilExcel.establecerTexto(r, 5, h.isNoturno() ? "Sí" : "No", null);        // HORARIO NOTURNO
                        UtilExcel.establecerTexto(r, 6, nvl(h.getDocumento()), null);              // DOCUMENTO
                        UtilExcel.establecerValor(r, 7, d.getOrden(), null);                       // ORDEN
                        UtilExcel.establecerTexto(r, 8, nvl(d.getHora()), null);                   // HORA
                        UtilExcel.establecerTexto(r, 9, nvl(d.getTolerancia()), null);             // TOLERANCIA
                        UtilExcel.establecerTexto(r,10, nvl(d.getTipoAccionShow()), null);         // ACCIÓN
                        UtilExcel.establecerTexto(r,11, d.isSegundoDia() ? "Sí" : "No", null);     // OTRO DÍA
                        UtilExcel.establecerValor(r,12, d.getMinutosAntes(), null);                // MINUTOS ANTES
                        UtilExcel.establecerValor(r,13, d.getMinutosDespues(), null);              // MINUTOS DESPUÉS
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, encabezados.length - 1, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada con filtros (A6: Nn)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "HorariosTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[]{ false, true, true, true, true, true, true, true, true, true, true, true, true, true }
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
    //            CSV
    // =========================
    public byte[] generarReporteCSV(ReporteHorariosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados como en tu exportación legacy
            sb.append("n,horario,codigo,horas_trabajo,minutos_alimentacion,horario_noturno,documento,")
              .append("orden,hora,tolerancia,accion,otro_dia,minutos_antes,minutos_despues\n");

            int n = 1;
            List<HorarioDTO> horarios = request.getHorarios();
            if (horarios != null) {
                for (HorarioDTO h : horarios) {
                    List<DetalleHorarioDTO> dets = h.getDetalles();
                    if (dets == null || dets.isEmpty()) continue;

                    for (DetalleHorarioDTO d : dets) {
                        String horario            = nvl(h.getNombre());
                        String codigo             = nvl(h.getCodigo());
                        String horasTrabajo       = nvl(h.getHoraTrabajo());
                        String minutosAliment     = nvl(h.getMinutosComida());
                        String noturno            = h.isNoturno() ? "Sí" : "No";
                        String documento          = nvl(h.getDocumento());
                        String orden              = nvl(d.getOrden());
                        String hora               = nvl(d.getHora());
                        String tolerancia         = nvl(d.getTolerancia());
                        String accion             = nvl(d.getTipoAccionShow());
                        String otroDia            = d.isSegundoDia() ? "Sí" : "No";
                        String minutosAntes       = nvl(d.getMinutosAntes());
                        String minutosDespues     = nvl(d.getMinutosDespues());

                        sb.append(n++).append(',')
                          .append(csv(horario)).append(',')
                          .append(csv(codigo)).append(',')
                          .append(csv(horasTrabajo)).append(',')
                          .append(csv(minutosAliment)).append(',')
                          .append(csv(noturno)).append(',')
                          .append(csv(documento)).append(',')
                          .append(csv(orden)).append(',')
                          .append(csv(hora)).append(',')
                          .append(csv(tolerancia)).append(',')
                          .append(csv(accion)).append(',')
                          .append(csv(otroDia)).append(',')
                          .append(csv(minutosAntes)).append(',')
                          .append(csv(minutosDespues)).append('\n');
                    }
                }
            }

            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //            XML
    // =========================
    public byte[] generarReporteXML(ReporteHorariosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Horarios>\n");

            List<HorarioDTO> horarios = request.getHorarios();
            if (horarios != null) {
                for (HorarioDTO h : horarios) {
                    sb.append("  <horario codigo=\"").append(xml(h.getCodigo())).append("\">\n");
                    sb.append("    <nombre>").append(xml(h.getNombre())).append("</nombre>\n");
                    sb.append("    <horas_trabajo>").append(xml(h.getHoraTrabajo())).append("</horas_trabajo>\n");
                    sb.append("    <minutos_alimentación>").append(xml(h.getMinutosComida())).append("</minutos_alimentación>\n");
                    sb.append("    <horario_noturno>").append(h.isNoturno() ? "Sí" : "No").append("</horario_noturno>\n");
                    sb.append("    <documento>").append(xml(h.getDocumento())).append("</documento>\n");

                    List<DetalleHorarioDTO> dets = h.getDetalles();
                    if (dets != null && !dets.isEmpty()) {
                        sb.append("    <detalles>\n");
                        for (DetalleHorarioDTO d : dets) {
                            sb.append("      <detalle orden=\"").append(xml(d.getOrden())).append("\">\n");
                            sb.append("        <hora>").append(xml(d.getHora())).append("</hora>\n");
                            sb.append("        <tolerancia>").append(xml(d.getTolerancia())).append("</tolerancia>\n");
                            sb.append("        <accion>").append(xml(d.getTipoAccionShow())).append("</accion>\n");
                            sb.append("        <otro_dia>").append(d.isSegundoDia() ? "Sí" : "No").append("</otro_dia>\n");
                            sb.append("        <minutos_antes>").append(xml(d.getMinutosAntes())).append("</minutos_antes>\n");
                            sb.append("        <minutos_despues>").append(xml(d.getMinutosDespues())).append("</minutos_despues>\n");
                            sb.append("      </detalle>\n");
                        }
                        sb.append("    </detalles>\n");
                    }

                    sb.append("  </horario>\n");
                }
            }

            sb.append("</Horarios>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //        Helpers locales
    // =========================
    private String nvl(Object v) {
        return v == null ? "" : String.valueOf(v);
    }

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

