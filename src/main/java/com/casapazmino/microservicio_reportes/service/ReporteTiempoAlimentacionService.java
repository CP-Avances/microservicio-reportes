package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTiempoAlimentacionService {

    public byte[] generarReporteTiempoAlimentacionPDF(ReporteTiempoAlimentacionRequest request) {
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

            String titulo = "TIEMPO DE ALIMENTACIÓN - "
                    + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));

            document.add(ReporteUtil
                    .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();
            Color colorExceso = new Color(0x55EE44);
            Color colorFT = new Color(0xEE4444);

            PdfPTable colores = new PdfPTable(5);
            colores.setWidthPercentage(100);
            colores.setWidths(new float[] { 3, 1.5f, 1.5f, 2, 2 });

            colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0xEE4444)));
            colores.addCell(ReporteUtil.celdaEncabezado("EXCESO DE ALIMENTACIÓN", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0x55EE44)));

            colores.setSpacingAfter(10f);
            document.add(colores);

            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados()
                            .forEach(emp -> contadorGlobal.addAndGet(emp.getAlimentacion().size())));

            PdfPTable tablaTitulo = new PdfPTable(2);
            tablaTitulo.setWidthPercentage(100);
            tablaTitulo.setWidths(new float[] { 8, 2 });

            PdfPCell celda1 = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celda1.setBackgroundColor(colorSecundario);
            celda1.setPadding(5f);
            celda1.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tablaTitulo.addCell(celda1);

            PdfPCell celda2 = new PdfPCell(
                    new Phrase("Nº Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celda2.setBackgroundColor(colorSecundario);
            celda2.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celda2.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celda2.setPadding(5);
            celda2.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tablaTitulo.addCell(celda2);

            tablaTitulo.setSpacingAfter(10f);
            document.add(tablaTitulo);

            for (GrupoAlimentacionDTO grupo : request.getGrupos()) {
                for (EmpleadoAlimentacionDTO emp : grupo.getEmpleados()) {

                    double totalExcesoAlimentacion = 0;
                    int contador = 1;

                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                            emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
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
                    tablaContenedora.setSpacingAfter(3f);
                    document.add(tablaContenedora);

                    PdfPTable tablaAlimentacion = new PdfPTable(7);
                    tablaAlimentacion.setWidthPercentage(100);
                    tablaAlimentacion.setWidths(new float[] { 0.5f, 1.8f, 1.8f, 1.8f, 1.8f, 1.8f, 1.8f });

                    String[] headers = { "Nº", "FECHA", "INICIO ALIMENTACIÓN", "FIN ALIMENTACIÓN", "M. ALIMENTACIÓN",
                            "M. TOMADOS", "M. EXCESO" };
                    for (String h : headers) {
                        tablaAlimentacion
                                .addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    }

                    for (RegistroAlimentacionDTO registro : emp.getAlimentacion()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                        String inicio = (registro.getInicioAlimentacion() == null
                                || registro.getInicioAlimentacion().isEmpty()) ? "FT"
                                        : extraerHora(registro.getInicioAlimentacion());
                        String fin = (registro.getFinAlimentacion() == null || registro.getFinAlimentacion().isEmpty())
                                ? "FT"
                                : extraerHora(registro.getFinAlimentacion());

                        Color colorInicio = "FT".equals(inicio) ? colorFT : fondo;
                        Color colorFin = "FT".equals(fin) ? colorFT : fondo;

                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(String.valueOf(contador), fondo));
                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(registro.getFecha(), fondo));
                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(inicio, colorInicio));
                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(fin, colorFin));
                        tablaAlimentacion.addCell(
                                ReporteUtil.celdaCentro(String.valueOf((int) registro.getMinutosPermitidos()), fondo));
                        tablaAlimentacion.addCell(
                                ReporteUtil.celdaCentro(String.valueOf(registro.getMinutosTomados()).replace(",", "."),
                                        fondo));
                        Color fondoExceso = registro.getMinutosExceso() > 0 ? colorExceso : fondo;
                        tablaAlimentacion.addCell(
                                ReporteUtil.celdaCentro(
                                        String.format("%.2f", registro.getMinutosExceso()).replace(",", "."),
                                        fondoExceso));

                        totalExcesoAlimentacion += registro.getMinutosExceso();
                        contador++;
                    }

                    for (int i = 0; i < 5; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER);
                        tablaAlimentacion.addCell(celdaVacia);
                    }
                    tablaAlimentacion.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaAlimentacion.addCell(
                            ReporteUtil.crearCelda(String.format("%.2f", totalExcesoAlimentacion).replace(",", "."),
                                    ReporteUtil.fuenteTexto(), Color.WHITE));

                    tablaAlimentacion.setSpacingAfter(10f);
                    document.add(tablaAlimentacion);
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
            return fechaHora;
        return fechaHora.split(" ")[1];
    }

}
