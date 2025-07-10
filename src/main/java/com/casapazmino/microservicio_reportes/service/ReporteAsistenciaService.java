package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ResumenAsistencia.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAsistenciaService {

    public byte[] generarReporteResumenAsistenciaPDF(ReporteAsistenciaRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            // Logo y encabezados
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            String titulo = "RESUMEN DE ASISTENCIA - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil
                    .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Tabla de codigos de color
            PdfPTable colores = new PdfPTable(6);
            colores.setWidthPercentage(100);
            colores.setWidths(new float[] { 2, 2, 2, 2, 2, 2 });

            colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", new Color(0xEE4444)));
            colores.addCell(ReporteUtil.celdaEncabezado("ATRASO", new Color(0xEEE344)));
            colores.addCell(ReporteUtil.celdaEncabezado("SALIDA ANTICIPADA", new Color(0x4499EE)));
            colores.addCell(ReporteUtil.celdaEncabezado("EXCESO DE ALIMENTACIÓN", new Color(0x55EE44)));
            colores.addCell(ReporteUtil.celdaEncabezado("VACACIONES", new Color(0xE68A2E)));

            colores.setSpacingAfter(10f);
            document.add(colores);

            // Contador global
            AtomicInteger totalRegistros = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados().forEach(emp -> totalRegistros.addAndGet(emp.getTLaborado().size())));
 
            //Tabla n registros
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(new float[] { 8, 2 });
            tituloTabla.setSpacingAfter(5f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + totalRegistros.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);
        
            tituloTabla.setSpacingAfter(10f);
            document.add(tituloTabla);


            for (GrupoAsistenciaDTO grupo : request.getGrupos()) {
                
                for (EmpleadoAsistenciaDTO emp : grupo.getEmpleados()) {
                    double totalAtrasos = 0;
                    double totalSalidasAnticipadas = 0;
                    double totalAlimentacionTomado = 0;
                    double totalAlimentacionAsignado = 0;
                    double totalLaborado = 0;
                    
                    //TABLA INFORMACION DEL EMPLEADO    
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
                    tablaContenedora.setSpacingAfter(5f);
                    document.add(tablaContenedora);


                    // TABLA DE ASISTENCIA POR DÍA (encabezado en una sola tabla con 16 columnas)
                    PdfPTable encabezado = new PdfPTable(16);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(new float[] {0.5f, 1.8f,
                            1.2f, 1.2f, 1.2f, 1.2f,
                            1.2f, 1.2f, 1.2f, 1.2f,
                            1.5f, 1.5f, 1.2f, 1.2f,
                            1.5f, 3f });
                    // Fila 1 - encabezados agrupados con rowspan o colspan
                    encabezado.addCell(
                            ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                    encabezado.addCell(
                            ReporteUtil.crearCelda("ENTRADA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("INICIO ALIMENTACIÓN", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("FIN ALIMENTACIÓN", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 1, 2));
                    encabezado.addCell(
                            ReporteUtil.crearCelda("SALIDA", ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));

                    encabezado.addCell(
                            ReporteUtil.crearCelda("ATRASO", ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                    encabezado.addCell(ReporteUtil.crearCelda("SALIDA ANTICIPADA", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 2, 1));

                    encabezado.addCell(ReporteUtil.crearCelda("T. ALIMENTACIÓN", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 1, 2));
                    encabezado.addCell(ReporteUtil.crearCelda("TIEMPO LABORADO", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 2, 1));
                    encabezado.addCell(ReporteUtil.crearCelda("OBSERVACIONES", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal, 2, 1));

                    // Fila 2 - subcolumnas debajo de agrupados
                    for (int i = 0; i < 4; i++) {
                        encabezado.addCell(
                                ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                        encabezado.addCell(
                                ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorSecundario));
                    }
                    encabezado.addCell(
                            ReporteUtil.crearCelda("ASIGNADO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    encabezado
                            .addCell(ReporteUtil.crearCelda("TOMADO", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    encabezado.setSpacingAfter(0f);        
                    document.add(encabezado);

                    // TABLA DE DATOS
                    PdfPTable tablaData = new PdfPTable(16);
                    tablaData.setWidthPercentage(100);
                    tablaData.setWidths(new float[] {0.5f, 1.8f,
                            1.2f, 1.2f, 1.2f, 1.2f,
                            1.2f, 1.2f, 1.2f, 1.2f,
                            1.5f, 1.5f, 1.2f, 1.2f,
                            1.5f, 3f });

                    Color colorFT = new Color(0xEE4444);
                    Color colorAtraso = new Color(0xEEE344);
                    Color colorSalidaAnticipada = new Color(0x4499EE);
                    Color colorExcesoAlimentacion= new Color(0x55EE44);

                    int contador = 1;
                    for (RegistroAsistenciaDTO reg : emp.getTLaborado()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                        tablaData.addCell(
                                ReporteUtil.crearCelda(String.valueOf(contador), ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda(
                                ReporteUtil.formatearFechaConDia(reg.getEntrada().getFecha_horario()),
                                ReporteUtil.fuenteTexto(), fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(extraerHora(reg.getEntrada().getFecha_hora_horario()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(
                                        ReporteUtil.crearCelda(
                                                formatearTimbre(reg.getEntrada().getFecha_hora_horario(),
                                                        reg.getEntrada().getFecha_hora_timbre()),
                                                ReporteUtil.fuenteTexto(),
                                                getColorTimbre(
                                                        formatearTimbre(reg.getEntrada().getFecha_hora_horario(),
                                                                reg.getEntrada().getFecha_hora_timbre()),
                                                        fondo, colorFT)));

                        tablaData.addCell(
                                ReporteUtil.crearCelda(extraerHora(reg.getInicioAlimentacion().getFecha_hora_horario()),
                                        ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(ReporteUtil.crearCelda(
                                        formatearTimbre(reg.getInicioAlimentacion().getFecha_hora_horario(),
                                                reg.getInicioAlimentacion().getFecha_hora_timbre()),
                                        ReporteUtil.fuenteTexto(),
                                        getColorTimbre(
                                                formatearTimbre(reg.getInicioAlimentacion().getFecha_hora_horario(),
                                                        reg.getInicioAlimentacion().getFecha_hora_timbre()),
                                                fondo, colorFT)));

                        tablaData.addCell(
                                ReporteUtil.crearCelda(extraerHora(reg.getFinAlimentacion().getFecha_hora_horario()),
                                        ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(ReporteUtil.crearCelda(
                                        formatearTimbre(reg.getFinAlimentacion().getFecha_hora_horario(),
                                                reg.getFinAlimentacion().getFecha_hora_timbre()),
                                        ReporteUtil.fuenteTexto(),
                                        getColorTimbre(
                                                formatearTimbre(reg.getFinAlimentacion().getFecha_hora_horario(),
                                                        reg.getFinAlimentacion().getFecha_hora_timbre()),
                                                fondo, colorFT)));

                        tablaData.addCell(ReporteUtil.crearCelda(extraerHora(reg.getSalida().getFecha_hora_horario()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData
                                .addCell(
                                        ReporteUtil.crearCelda(
                                                formatearTimbre(reg.getSalida().getFecha_hora_horario(),
                                                        reg.getSalida().getFecha_hora_timbre()),
                                                ReporteUtil.fuenteTexto(),
                                                getColorTimbre(
                                                        formatearTimbre(reg.getSalida().getFecha_hora_horario(),
                                                                reg.getSalida().getFecha_hora_timbre()),
                                                        fondo, colorFT)));

                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(reg.getMinAtrasos()),
                                ReporteUtil.fuenteTexto(), reg.getMinAtrasos() > 0 ? colorAtraso : fondo));
                                
                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getMinSalidasAnticipadas()), ReporteUtil.fuenteTexto(),
                                reg.getMinSalidasAnticipadas() > 0 ? colorSalidaAnticipada : fondo));

                        tablaData.addCell(ReporteUtil.crearCelda(
                                convertirMinutosATiempo(reg.getInicioAlimentacion().getMinutos_alimentacion()),
                                ReporteUtil.fuenteTexto(), fondo));


                        tablaData.addCell(ReporteUtil.crearCelda(
                        convertirMinutosATiempo(reg.getMinAlimentacion()),
                        ReporteUtil.fuenteTexto(),
                        (reg.getInicioAlimentacion().getMinutos_alimentacion() != null &&
                        reg.getMinAlimentacion() != null &&
                        reg.getMinAlimentacion() > reg.getInicioAlimentacion().getMinutos_alimentacion())
                                ? colorExcesoAlimentacion
                                : fondo
                        ));

                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(reg.getMinLaborados()),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), fondo));

                        contador++;

                        totalAtrasos += reg.getMinAtrasos() != null ? reg.getMinAtrasos() : 0;
                        totalSalidasAnticipadas += reg.getMinSalidasAnticipadas() != null ? reg.getMinSalidasAnticipadas() : 0;
                        totalAlimentacionTomado += reg.getMinAlimentacion() != null ? reg.getMinAlimentacion() : 0;
                        totalAlimentacionAsignado += reg.getInicioAlimentacion().getMinutos_alimentacion() != null ? reg.getInicioAlimentacion().getMinutos_alimentacion() : 0;
                        totalLaborado += reg.getMinLaborados() != null ? reg.getMinLaborados() : 0;
                    }

                    // Fila final de totales
                        Color fondoTotal = new Color(230, 240, 255);

                        // Vacías hasta columna 10 (horarios y timbres)
                        for (int i = 0; i < 9; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER); // elimina todos los bordes
                        tablaData.addCell(celdaVacia);
                        }
                        tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE)); // Fecha

                        // Totales
                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalAtrasos), ReporteUtil.fuenteTexto(), Color.WHITE));  // Atrasos
                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalSalidasAnticipadas), ReporteUtil.fuenteTexto(), Color.WHITE)); // Salida anticipada
                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalAlimentacionAsignado), ReporteUtil.fuenteTexto(), Color.WHITE)); // Asignado
                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalAlimentacionTomado), ReporteUtil.fuenteTexto(), Color.WHITE)); // Tomado
                        tablaData.addCell(ReporteUtil.crearCelda(convertirMinutosATiempo(totalLaborado), ReporteUtil.fuenteTexto(), Color.WHITE)); // Laborado
                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE)); // Observaciones

                    tablaData.setSpacingAfter(10f);  
                    document.add(tablaData);
                    

                }
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private String convertirMinutosATiempo(Double minutos) {
        if (minutos == null || minutos <= 0)
            return "00:00:00";
        int totalSegundos = (int) Math.round(minutos * 60);
        int horas = totalSegundos / 3600;
        int mins = (totalSegundos % 3600) / 60;
        int segundos = totalSegundos % 60;
        return String.format("%02d:%02d:%02d", horas, mins, segundos);
    }

    private String extraerHora(String fechaHora) {
        if (fechaHora == null || !fechaHora.contains(" "))
            return "";
        return fechaHora.split(" ")[1];
    }

    private String formatearTimbre(String horario, String timbre) {
        String horaHorario = extraerHora(horario);

        if (horaHorario == null || horaHorario.isEmpty()) {
            return "";
        }

        if (timbre == null || timbre.trim().isEmpty()) {
            if ("00:00:00".equals(horaHorario) || "23:59:00".equals(horaHorario)) {
                return "L";
            }
            return "FT";
        }

        return extraerHora(timbre);
    }

    private Color getColorTimbre(String valor, Color porDefecto, Color colorFT) {
        if ("FT".equals(valor))
            return colorFT;
        return porDefecto;
    }

}
