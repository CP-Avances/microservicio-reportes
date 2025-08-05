package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteFaltas.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteFaltasService {

    public byte[] generarReporteFaltasPDF(ReporteFaltasRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4, 40, 40, 30, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(
                    "FALTAS - USUARIOS " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS")));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Contador global
            AtomicInteger totalFaltasGeneral = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados()
                            .forEach(emp -> totalFaltasGeneral.addAndGet(emp.getFaltas().size()))
            );

            // Título con contador global
            PdfPTable tablaTitulo = new PdfPTable(2);
            tablaTitulo.setWidthPercentage(100);
            tablaTitulo.setWidths(new float[]{8, 2});

            PdfPCell celda1 = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celda1.setBackgroundColor(colorSecundario);
            celda1.setPadding(5f);
            celda1.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tablaTitulo.addCell(celda1);

            PdfPCell celda2 = new PdfPCell(new Phrase("Nº Registros: " + totalFaltasGeneral.get(), ReporteUtil.fuenteEncabezado()));
            celda2.setBackgroundColor(colorSecundario);
            celda2.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celda2.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celda2.setPadding(5f);
            celda2.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tablaTitulo.addCell(celda2);

            tablaTitulo.setSpacingAfter(10f);
            document.add(tablaTitulo);

            // Por grupo
            for (GrupoFaltasDTO grupo : request.getGrupos()) {
                for (EmpleadoFaltasDTO emp : grupo.getEmpleados()) {
                    int contador = 1;

                    // Tabla info empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[]{4, 4, 4});

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));


                    // Contenedor con borde
                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    tablaContenedora.setSpacingAfter(3f);
                    document.add(tablaContenedora);

                    // Tabla faltas
                    PdfPTable tablaFaltas = new PdfPTable(2);
                    tablaFaltas.setWidthPercentage(100);
                    tablaFaltas.setWidths(new float[]{3, 3});

                    tablaFaltas.addCell(ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaFaltas.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    for (FaltaDTO falta : emp.getFaltas()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;
                        tablaFaltas.addCell(ReporteUtil.celdaCentro(String.valueOf(contador), fondo));
                        tablaFaltas.addCell(ReporteUtil.celdaCentro(ReporteUtil.formatearFechaConDia(falta.getFecha()), fondo));
                        contador++;
                    }

                    tablaFaltas.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), colorSecundario));
                    tablaFaltas.addCell(ReporteUtil.crearCelda(String.valueOf(emp.getFaltas().size()),
                            ReporteUtil.fuenteTexto(), colorSecundario));

                    tablaFaltas.setSpacingAfter(10f);
                    document.add(tablaFaltas);
                }
            }

            // Resumen general final
            if (!request.isResumen()) {
                document.add(Chunk.NEWLINE);
                PdfPTable resumen = new PdfPTable(3);
                resumen.setWidthPercentage(100);
                resumen.setWidths(new float[]{4, 4, 2});
                resumen.setSpacingBefore(10f);

                resumen.addCell(ReporteUtil.crearCelda("TOTAL GENERAL", ReporteUtil.fuenteEncabezado(), colorSecundario));
                resumen.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorSecundario));
                resumen.addCell(ReporteUtil.crearCelda(String.valueOf(totalFaltasGeneral.get()),
                        ReporteUtil.fuenteEncabezado(), colorSecundario));

                document.add(resumen);
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
