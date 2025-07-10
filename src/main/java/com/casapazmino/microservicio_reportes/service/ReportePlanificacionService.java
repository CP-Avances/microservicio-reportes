package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionDetalleDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionEmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionHorarioMensualDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.ReportePlanificacionRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

@Service
public class ReportePlanificacionService {

    public byte[] generarReportePDF(ReportePlanificacionRequest request) {
        try {
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            ConfiguracionPaginaPDF evento = new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal());
            writer.setPageEvent(evento);
            document.open();

            if (request.getLogoBase64() != null) {
                Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                document.add(logo);
            }

            Paragraph nombreEmpresa = new Paragraph(request.getEmpresa(), ReporteUtil.fuenteEncabezado());
            nombreEmpresa.setAlignment(Element.ALIGN_CENTER);
            document.add(nombreEmpresa);

            Paragraph tituloPrincipal = new Paragraph(request.getTitulo(), ReporteUtil.fuenteEncabezado());
            tituloPrincipal.setAlignment(Element.ALIGN_CENTER);
            document.add(tituloPrincipal);

            Paragraph periodo = new Paragraph(
                    "PERIODO DEL: " + request.getPeriodoInicio() + " AL " + request.getPeriodoFin(),
                    ReporteUtil.fuenteEncabezado());
            periodo.setAlignment(Element.ALIGN_CENTER);
            document.add(periodo);
            document.add(Chunk.NEWLINE);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

            // TABLA DE HORARIOS
            PdfPTable tablaHorarios = new PdfPTable(5);
            tablaHorarios.setWidthPercentage(100f);
            tablaHorarios.setHorizontalAlignment(Element.ALIGN_LEFT);
            tablaHorarios.setSpacingAfter(20f);
            tablaHorarios.setWidths(new int[] { 16, 16, 16, 16, 16 });

            PdfPCell tituloHorarios = new PdfPCell(new Phrase("DETALLE DE HORARIOS", ReporteUtil.fuenteEncabezado()));
            tituloHorarios.setColspan(5);
            tituloHorarios.setHorizontalAlignment(Element.ALIGN_CENTER);
            tituloHorarios.setBackgroundColor(colorPrincipal);
            tablaHorarios.addCell(tituloHorarios);

            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("HORARIO", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("ENTRADA (E)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("INICIO ALIMENTACIÓN (I/A)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("FIN ALIMENTACIÓN (F/A)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("SALIDA (S)", colorPrincipal));

            for (PlanificacionDetalleDTO d : request.getDetalle_acciones()) {
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getHorario()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getEntrada_()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getInicio_comida()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getFin_comida()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getSalida_()));
            }

            // TABLA DE NOMENCLATURA
            PdfPTable tablaNomenclatura = new PdfPTable(2);
            tablaNomenclatura.setWidthPercentage(80f);
            tablaNomenclatura.setHorizontalAlignment(Element.ALIGN_RIGHT);
            tablaNomenclatura.setSpacingAfter(100f);
            tablaNomenclatura.setWidths(new int[] { 5, 7 });

            PdfPCell tituloNomen = new PdfPCell(new Phrase("DEFINICIONES", ReporteUtil.fuenteEncabezado()));
            tituloNomen.setColspan(2);
            tituloNomen.setHorizontalAlignment(Element.ALIGN_CENTER);
            tituloNomen.setBackgroundColor(colorSecundario);
            tablaNomenclatura.addCell(tituloNomen);

            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("NOMENCLATURA", colorSecundario));
            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("DESCRIPCIÓN", colorSecundario));

            for (Map<String, String> def : request.getNomenclatura()) {
                tablaNomenclatura.addCell(ReporteUtil.celdaNomenclatura(def.get("nombre"), ReporteUtil.fuenteTexto()));
                tablaNomenclatura.addCell(
                        ReporteUtil.celdaNomenclaturaDescripcion(def.get("descripcion"), ReporteUtil.fuenteTexto()));
            }

            // Generar primero la tablaNomenclatura completamente para medir su altura real
            tablaNomenclatura.setKeepTogether(true);
            tablaNomenclatura.completeRow(); // asegura filas completas
            tablaNomenclatura.calculateHeights(true); // o false si ya estaba calculada

            // Tabla contenedora de una fila con dos celdas independientes
            PdfPTable tablaContenedora = new PdfPTable(2);
            tablaContenedora.setWidthPercentage(70f); // reduce el espacio total ocupado (de 100% a 70%)
            tablaContenedora.setWidths(new float[] { 60, 40 }); // proporción interna: 60% horarios, 40% definiciones
            tablaContenedora.setSpacingBefore(5f); // opcional: reducir espacio antes

            // Celda de la tabla de horarios (puede crecer)
            PdfPCell celdaIzquierda = new PdfPCell();
            celdaIzquierda.setBorder(Rectangle.NO_BORDER);
            celdaIzquierda.setVerticalAlignment(Element.ALIGN_TOP);
            celdaIzquierda.addElement(tablaHorarios);
            tablaContenedora.addCell(celdaIzquierda);

            // Celda de la tabla de definiciones (no se estira)
            PdfPCell celdaDerecha = new PdfPCell();
            celdaDerecha.setBorder(Rectangle.NO_BORDER);
            celdaDerecha.setVerticalAlignment(Element.ALIGN_TOP);
            celdaDerecha.addElement(tablaNomenclatura);
            tablaContenedora.addCell(celdaDerecha);

            // Agregar al documento
            document.add(tablaContenedora);

            for (PlanificacionEmpleadoDTO emp : request.getDatos()) {
                PdfPTable encabezado = new PdfPTable(3);
                encabezado.setWidthPercentage(100);
                encabezado.setWidths(new int[] { 5, 3, 3 });

                encabezado.addCell(
                        ReporteUtil.celdaInfoEmpleado("EMPLEADO: " + emp.getApellido() + " " + emp.getNombre()));
                encabezado.addCell(ReporteUtil.celdaInfoEmpleado("C.C.: " + emp.getIdentificacion()));
                encabezado.addCell(ReporteUtil.celdaInfoEmpleado("COD: " + emp.getCodigo()));

                encabezado.addCell(ReporteUtil.celdaInfoEmpleado("DEPARTAMENTO: " + emp.getDepartamento()));
                encabezado.addCell(ReporteUtil.celdaInfoEmpleado("CARGO: " + emp.getCargo()));
                encabezado.addCell(ReporteUtil.celdaInfoEmpleado(""));
                document.add(encabezado);
                Paragraph espacio = new Paragraph("", new Font());
                espacio.setSpacingBefore(10f); // o el valor que necesites
                document.add(espacio);

                for (PlanificacionHorarioMensualDTO mes : emp.getHorarios()) {
                    PdfPTable tablaMes = new PdfPTable(7);
                    tablaMes.setWidthPercentage(100);

                    PdfPCell celdaTituloMes = new PdfPCell(new Phrase("AÑO: " + mes.getAnio() + " MES: " + mes.getMes(),
                            ReporteUtil.fuenteEncabezado()));
                    celdaTituloMes.setColspan(7);
                    celdaTituloMes.setBackgroundColor(colorSecundario);
                    celdaTituloMes.setHorizontalAlignment(Element.ALIGN_CENTER);
                    tablaMes.addCell(celdaTituloMes);

                    for (int i = 1; i <= 31; i += 7) {
                        for (int j = i; j < i + 7; j++) {
                            if (j <= 31) {
                                PdfPCell celdaDia = new PdfPCell(
                                        new Phrase(String.format("%02d", j), ReporteUtil.fuenteEncabezado()));
                                celdaDia.setBackgroundColor(colorPrincipal); // o colorSecundario si prefieres
                                celdaDia.setHorizontalAlignment(Element.ALIGN_CENTER);
                                celdaDia.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                tablaMes.addCell(celdaDia);
                            } else {
                                tablaMes.addCell("");
                            }
                        }

                        for (int j = i; j < i + 7; j++) {
                            if (j <= 31) {
                                Method getter = PlanificacionHorarioMensualDTO.class.getMethod("getDia" + j);
                                String valor = (String) getter.invoke(mes);
                                tablaMes.addCell(ReporteUtil.celdaCentro(valor != null ? valor : ""));
                            } else {
                                tablaMes.addCell("");
                            }
                        }
                    }

                    document.add(tablaMes);
                    Paragraph espacioEntreEmpleados = new Paragraph("", new Font());
                    espacioEntreEmpleados.setSpacingBefore(25f); 
                    document.add(espacioEntreEmpleados);

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