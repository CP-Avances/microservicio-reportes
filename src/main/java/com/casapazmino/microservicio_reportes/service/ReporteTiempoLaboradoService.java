package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTiempoLaboradoService {

        public byte[] generarReporteTiempoLaboradoPDF(ReporteTiempoLaboradoRequest request) {
                try {
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        Document document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
                        PdfWriter writer = PdfWriter.getInstance(document, baos);
                        writer.setPageEvent(new ConfiguracionPaginaPDF(
                                        request.getUsuario(),
                                        request.getFraseMarcaAgua(),
                                        request.getColorPrincipal()));

                        document.open();

                        // Logo y encabezado
                        Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                        if (logo != null)
                                document.add(logo);

                        document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
                        String titulo = "REPORTE DE TIEMPO LABORADO - "
                                        + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
                        document.add(ReporteUtil.crearTituloReporte(titulo));
                        document.add(ReporteUtil
                                        .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL "
                                                        + request.getFechaFin()));

                        Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
                        Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
                        Color zebraColor = ReporteUtil.colorZebraClaro();

                        // Contador
                        AtomicInteger totalRegistros = new AtomicInteger();
                        request.getGrupos().forEach(
                                        grupo -> grupo.getEmpleados().forEach(emp -> {
                                                totalRegistros.addAndGet(emp.getTLaborado().size());
                                        }));

                        // Tabla de codigos de color
                        PdfPTable colores = new PdfPTable(5);
                        colores.setWidthPercentage(100);
                        colores.setWidths(new float[] { 3, 1.5f, 1.5f, 2, 2 });

                        colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0xEE4444)));
                        colores.addCell(ReporteUtil.celdaEncabezado("TIEMPO LABORADO MENOR AL PLANIFICADO",
                                        Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0x55EE44)));

                        colores.setSpacingAfter(10f);
                        document.add(colores);

                        // Tabla de título y contador
                        PdfPTable tablaTitulo = new PdfPTable(2);
                        tablaTitulo.setWidthPercentage(100);
                        tablaTitulo.setWidths(new float[] { 8, 2 });
                        tablaTitulo.setSpacingAfter(10f);

                        PdfPCell celdaTitulo = new PdfPCell(
                                        new Phrase("LISTA DE EMPLEADOS", ReporteUtil.fuenteEncabezado()));
                        celdaTitulo.setBackgroundColor(colorSecundario);
                        celdaTitulo.setPadding(5f);
                        celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                        tablaTitulo.addCell(celdaTitulo);

                        PdfPCell celdaContador = new PdfPCell(
                                        new Phrase("N° Registros: " + totalRegistros.get(),
                                                        ReporteUtil.fuenteEncabezado()));
                        celdaContador.setBackgroundColor(colorSecundario);
                        celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
                        celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        celdaContador.setPadding(5f);
                        celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                        tablaTitulo.addCell(celdaContador);

                        document.add(tablaTitulo);

                        for (GrupoTiempoDTO grupo : request.getGrupos()) {
                                for (EmpleadoTiempoDTO emp : grupo.getEmpleados()) {

                                        double totalTiempoPlanificado = 0;
                                        double totalTiempoPlanMinutos = 0;
                                        double totalTiempoLaborado = 0;
                                        double totalTiempoLabMinutos = 0;

                                        // TABLA INFORMACION DEL EMPLEADO
                                        PdfPTable infoEmpleado = new PdfPTable(3);
                                        infoEmpleado.setWidthPercentage(100);
                                        infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",
                                                        emp.getIdentificacion(), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                                        emp.getApellido() + " " + emp.getNombre(), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(),
                                                        zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:",
                                                        emp.getRegimen(), zebraColor));
                                        infoEmpleado
                                                        .addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:",
                                                                        emp.getDepartamento(), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(),
                                                        zebraColor));

                                        PdfPTable tablaContenedora = new PdfPTable(1);
                                        tablaContenedora.setWidthPercentage(100);
                                        PdfPCell contenedor = new PdfPCell(infoEmpleado);
                                        contenedor.setPadding(0);
                                        contenedor.setBorder(Rectangle.BOX);
                                        tablaContenedora.addCell(contenedor);
                                        tablaContenedora.setSpacingAfter(5f);
                                        document.add(tablaContenedora);

                                        // TABLA DE ASISTENCIA POR DÍA (encabezado en una sola tabla con 16 columnas)
                                        PdfPTable encabezado = new PdfPTable(14);
                                        encabezado.setWidthPercentage(100);
                                        encabezado.setWidths(new float[] { 0.5f, 1.8f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.5f, 1.5f, 1.5f, 1.5f });
                                        // Fila 1 - encabezados agrupados con rowspan o colspan
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(),
                                                                        colorPrincipal, 2, 1));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(),
                                                                        colorPrincipal, 2, 1));

                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("ENTRADA",
                                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal,
                                                                        1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("INICIO ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("FIN ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 1, 2));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("SALIDA", ReporteUtil.fuenteEncabezado(),
                                                                        colorPrincipal, 1, 2));

                                        encabezado.addCell(ReporteUtil.crearCelda("TIEMPO PLANIFICADO",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 2, 2));

                                        encabezado.addCell(ReporteUtil.crearCelda("TIEMPO LABORADO",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 2, 2));

                                        // Fila 2 - subcolumnas debajo de agrupados
                                        for (int i = 0; i < 4; i++) {
                                                encabezado.addCell(
                                                                ReporteUtil.crearCelda("HORARIO",
                                                                                ReporteUtil.fuenteEncabezado(),
                                                                                colorPrincipal));
                                                encabezado.addCell(
                                                                ReporteUtil.crearCelda("TIMBRE",
                                                                                ReporteUtil.fuenteEncabezado(),
                                                                                colorSecundario));
                                        }

                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("MINUTOS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("HH:MM:SS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));

                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("MINUTOS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("HH:MM:SS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));

                                        encabezado.setSpacingAfter(0f);
                                        document.add(encabezado);

                                        // TABLA DE DATOS
                                        PdfPTable tablaData = new PdfPTable(14);
                                        tablaData.setWidthPercentage(100);
                                        tablaData.setWidths(new float[] { 0.5f, 1.8f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.5f, 1.5f, 1.5f, 1.5f });

                                        Color colorFT = new Color(0xEE4444);
                                        Color colorTiempoMenorPlani = new Color(0x55EE44);

                                        int contador = 1;
                                        double totalPlanificadosMin = 0;
                                        double totalLaboradosMin = 0;

                                        for (RegistroTiempoDTO reg : emp.getTLaborado()) {
                                                Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                                                tablaData.addCell(
                                                                ReporteUtil.crearCelda(String.valueOf(contador),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                ReporteUtil.formatearFechaConDia(
                                                                                reg.getEntrada().getFecha_horario()),
                                                                ReporteUtil.fuenteTexto(), fondo));

                                                // ENTRADA
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                extraerHora(reg.getEntrada().getFecha_hora_horario()),
                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getEntrada()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getEntrada().getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg.getEntrada()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getEntrada().getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // INICIO ALIMENTACIÓN
                                                tablaData.addCell(
                                                                ReporteUtil.crearCelda(extraerHora(reg
                                                                                .getInicioAlimentacion()
                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getInicioAlimentacion()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getInicioAlimentacion()
                                                                                                .getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg
                                                                                                .getInicioAlimentacion()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getInicioAlimentacion()
                                                                                                                .getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // FIN ALIMENTACIÓN
                                                tablaData.addCell(
                                                                ReporteUtil.crearCelda(extraerHora(reg
                                                                                .getFinAlimentacion()
                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getFinAlimentacion()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getFinAlimentacion()
                                                                                                .getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg.getFinAlimentacion()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getFinAlimentacion()
                                                                                                                .getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // SALIDA
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                extraerHora(reg.getSalida().getFecha_hora_horario()),
                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getSalida().getFecha_hora_horario(),
                                                                                reg.getSalida().getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg.getSalida()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getSalida().getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // TIEMPO PLANIFICADO
                                                double minPlanificado = Double.parseDouble(reg.getMinPlanificados());
                                                double minLaborado = Double.parseDouble(reg.getMinLaborados());
                                                Color fondoTiempoLaborado = minLaborado < minPlanificado ? colorTiempoMenorPlani : fondo;

                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getTiempoPlanificado(), fondo));
                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getMinPlanificados(), fondo));

                                                // TIEMPO LABORADO
                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getTiempoLaborado(), fondoTiempoLaborado));
                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getMinLaborados(), fondoTiempoLaborado));

                                                
                                                contador++;
                                                
                                                totalPlanificadosMin += Double.parseDouble(reg.getMinPlanificados());
                                                totalLaboradosMin += Double.parseDouble(reg.getMinLaborados());

                                        }

                                        // Vacías hasta columna 10
                                        for (int i = 0; i < 9; i++) {
                                                PdfPCell celdaVacia = ReporteUtil.crearCelda("",
                                                                ReporteUtil.fuenteTexto(), Color.WHITE);
                                                celdaVacia.setBorder(Rectangle.NO_BORDER);
                                                tablaData.addCell(celdaVacia);
                                        }

                                        // Celda: TOTAL (Texto)
                                        tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(),
                                                        Color.WHITE));

                                        // Totales de TIEMPO PLANIFICADO
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        convertirMinutosATiempo(totalPlanificadosMin),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        String.format("%.2f", totalPlanificadosMin).replace(",", "."),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));

                                        // Totales de TIEMPO LABORADO
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        convertirMinutosATiempo(totalLaboradosMin),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        String.format("%.2f", totalLaboradosMin).replace(",", "."),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));

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

        private String convertirMinutosATiempo(double minutos) {
                if (minutos <= 0)
                        return "00:00:00";

                int totalSegundos = (int) Math.round(minutos * 60);
                int horas = totalSegundos / 3600;
                int mins = (totalSegundos % 3600) / 60;
                int segundos = totalSegundos % 60;

                return String.format("%02d:%02d:%02d", horas, mins, segundos);
        }

}
