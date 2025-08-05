package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.ReporteSalidasAnticipadasRequest;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.GrupoSalidasDTO;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.EmpleadoSalidaDTO;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.SalidaDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteSalidasAnticipadasService {

    public byte[] generarReportePDF(ReporteSalidasAnticipadasRequest request) {
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
            String titulo = "SALIDAS ANTICIPADAS - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil
                    .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados().forEach(emp -> contadorGlobal.addAndGet(emp.getSalidas().size())));

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

            int contador = 1;
            for (GrupoSalidasDTO grupo : request.getGrupos()) {
                for (EmpleadoSalidaDTO emp : grupo.getEmpleados()) {

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

                    PdfPTable tablaSalidas = new PdfPTable(12);
                    tablaSalidas.setWidthPercentage(100);
                    tablaSalidas.setSpacingBefore(5f);
                    tablaSalidas.setWidths(new float[] { 0.5f, 1.8f, 1.5f, 1.8f, 2, 2, 2, 2, 2, 2, 2, 2 });

                    PdfPCell h1 = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h1.setRowspan(2);
                    tablaSalidas.addCell(h1);
                    PdfPCell h2 = ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h2.setColspan(2);
                    tablaSalidas.addCell(h2);
                    PdfPCell h3 = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h3.setColspan(2);
                    tablaSalidas.addCell(h3);
                    PdfPCell h4 = ReporteUtil.crearCelda("TIPO PERMISO", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal);
                    h4.setRowspan(2);
                    tablaSalidas.addCell(h4);
                    PdfPCell h5 = ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h5.setRowspan(2);
                    tablaSalidas.addCell(h5);
                    PdfPCell h6 = ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h6.setRowspan(2);
                    tablaSalidas.addCell(h6);
                    PdfPCell h7 = ReporteUtil.crearCelda("PERMISO", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h7.setColspan(2);
                    tablaSalidas.addCell(h7);
                    PdfPCell h8 = ReporteUtil.crearCelda("SALIDA ANTICIPADA", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal);
                    h8.setColspan(2);
                    tablaSalidas.addCell(h8);

                    // Segunda fila encabezados (8 columnas)
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    long totalSegundos = 0;
                    double totalMinutos = 0;

                    for (SalidaDTO s : emp.getSalidas()) {

                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;
                        String[] horaHorario = s.getFecha_hora_horario().split(" ");
                        String[] horaTimbre = s.getFecha_hora_timbre().split(" ");

                        long segundos = s.getDiferencia() != null ? Math.round(s.getDiferencia()) : 0;
                        long horas = segundos / 3600;
                        long minutos = (segundos % 3600) / 60;
                        long restoSeg = segundos % 60;
                        String tiempoFormateado = String.format("%02d:%02d:%02d", horas, minutos, restoSeg);

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(String.valueOf(contador), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(horaHorario[0]),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(horaHorario.length > 1 ? horaHorario[1] : "",
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(horaTimbre[0]),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(horaTimbre.length > 1 ? horaTimbre[1] : "",
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas
                                .addCell(ReporteUtil.crearCelda(s.getTipo_permiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(s.getDesde(), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(s.getHasta(), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(" ", ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(" ", ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas
                                .addCell(ReporteUtil.crearCelda(tiempoFormateado, ReporteUtil.fuenteTexto(), fondo)); // hh:mm:ss
                        tablaSalidas.addCell(ReporteUtil.crearCelda(String.format("%.2f", s.getDiferencia() / 60.0),
                                ReporteUtil.fuenteTexto(), fondo));
                        contador++;
                        totalSegundos += segundos;
                        totalMinutos += s.getDiferencia() != null ? s.getDiferencia() / 60.0 : 0;

                    }

                    // 9 celdas vacías
                    for (int i = 0; i < 9; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER);
                        tablaSalidas.addCell(celdaVacia);
                    }

                    // Celda "TOTAL"
                    tablaSalidas.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE));

                    // Tiempo total formateado
                    long horasT = totalSegundos / 3600;
                    long minutosT = (totalSegundos % 3600) / 60;
                    long segRest = totalSegundos % 60;
                    String tiempoTotal = String.format("%02d:%02d:%02d", horasT, minutosT, segRest);
                    tablaSalidas.addCell(ReporteUtil.crearCelda(tiempoTotal, ReporteUtil.fuenteTexto(), Color.WHITE));

                    // Total en minutos
                    tablaSalidas.addCell(ReporteUtil.crearCelda(String.format("%.2f", totalMinutos),
                            ReporteUtil.fuenteTexto(), Color.WHITE));

                    document.add(tablaSalidas);
                    document.add(Chunk.NEWLINE);
                }
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
