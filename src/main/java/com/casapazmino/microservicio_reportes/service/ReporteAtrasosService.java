package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteAtrasos.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAtrasosService {

    public byte[] generarReportePDF(ReporteAtrasosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4, 40, 40, 30, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            String titulo = "REPORTE DE ATRASOS - " + request.getOpcionBusqueda();
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil
                    .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados().forEach(emp -> contadorGlobal.addAndGet(emp.getAtrasos().size())));

            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(new float[] { 8, 2 });
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            for (GrupoAtrasoDTO grupo : request.getGrupos()) {
                for (EmpleadoAtrasoDTO emp : grupo.getEmpleados()) {

                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                            emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado
                            .addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    document.add(tablaContenedora);

                    // TABLA DE ASISTENCIA POR DÍA (encabezado en una sola tabla con 16 columnas)
                    PdfPTable encabezado = new PdfPTable(11);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(new float[] { 0.5f, 1.2f,
                            1.2f, 1.2f, 1.2f, 1.3f,
                            1, 1, 1.4f, 1.8f,
                            2.6f });
                    // Fila 1 - encabezados agrupados con rowspan o colspan
                    encabezado.addCell(
                            ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("TIPO PERMISO", ReporteUtil.fuenteEncabezado(),
                                    colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezado(),
                                    colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("PERMISO", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("TOLERANCIA", ReporteUtil.fuenteEncabezado(),
                                    colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("ATRASO", ReporteUtil.fuenteEncabezado(),
                                    colorPrincipal, 2, 1));

                    // Fila 2 - subcolumnas debajo de agrupados
                    for (int i = 0; i < 2; i++) {
                        encabezado.addCell(
                                ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                        encabezado.addCell(
                                ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorSecundario));
                    }
                    encabezado.setSpacingAfter(0f);
                    document.add(encabezado);

                    PdfPTable tablaData = new PdfPTable(13);
                    tablaData.setWidthPercentage(100);
                    tablaData.setWidths(new float[] { 0.5f, 1.2f,
                            1.2f, 1.2f, 1.2f, 1.3f,
                            1, 1, 0.7f, 0.7f, 1.8f,
                            1.3F, 1.3F });

                    int contador = 1;
                    long totalSegundos = 0;
                    double totalMinutos = 0;

                    for (AtrasoDTO atraso : emp.getAtrasos()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                        tablaData.addCell(
                                ReporteUtil.crearCelda(String.valueOf(contador), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(atraso.getFechaHorario()),
                                        ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getHoraHorario(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(atraso.getFechaTimbre()),
                                        ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getHoraTimbre(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getTipo_permiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(atraso.getDesde(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(atraso.getHasta(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(ReporteUtil.crearCelda(atraso.getPermiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(ReporteUtil.crearCelda(atraso.getPermiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getTolerancia(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getTiempoAtraso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(
                                ReporteUtil.crearCelda(atraso.getMinutosAtraso(), ReporteUtil.fuenteTexto(), fondo));

                        contador++;
                        totalSegundos += convertirTiempoAtrasoASegundos(atraso.getTiempoAtraso());
                        totalMinutos += convertirMinutos(atraso.getMinutosAtraso());
                    }

                    // === Fila TOTAL al final de la tabla ===
                    Color fondoTotal = new Color(230, 240, 255);

                    // Celdas vacías hasta columna 9
                    for (int i = 0; i < 9; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER);
                        tablaData.addCell(celdaVacia);
                    }

                    // Texto TOTAL
                    tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), fondoTotal));

                    // Tiempo total formateado
                    long h = totalSegundos / 3600;
                    long m = (totalSegundos % 3600) / 60;
                    long s = totalSegundos % 60;
                    String totalFormateado = String.format("%02d:%02d:%02d", h, m, s);

                    // Tiempo atraso total y minutos
                    tablaData.addCell(ReporteUtil.crearCelda(totalFormateado, ReporteUtil.fuenteTexto(), fondoTotal));
                    tablaData.addCell(ReporteUtil.crearCelda(
                            String.format("%.2f", totalMinutos).replace(",", "."),
                            ReporteUtil.fuenteTexto(), fondoTotal));

                    // Agregar tabla al documento
                    tablaData.setSpacingAfter(10f);
                    document.add(tablaData);

                }

            }

            document.close();
            return baos.toByteArray();

        } catch (

        Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static long convertirTiempoAtrasoASegundos(String tiempo) {
        if (tiempo == null || !tiempo.matches("\\d{2}:\\d{2}:\\d{2}"))
            return 0;
        try {
            String[] partes = tiempo.split(":");
            int h = Integer.parseInt(partes[0]);
            int m = Integer.parseInt(partes[1]);
            int s = Integer.parseInt(partes[2]);
            return h * 3600 + m * 60 + s;
        } catch (Exception e) {
            return 0;
        }
    }

    public static double convertirMinutos(String minutosStr) {
        if (minutosStr == null)
            return 0;
        try {
            return Double.parseDouble(minutosStr.replace(",", "."));
        } catch (Exception e) {
            return 0;
        }
    }

}
