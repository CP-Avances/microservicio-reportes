package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.*;
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
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 50, 50);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            ConfiguracionPaginaPDF evento = new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal());
            writer.setPageEvent(evento);
            document.open();

            // LOGO + EMPRESA + TÍTULO
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

            // TABLA DE HORARIOS (se inserta dentro de una celda)
            PdfPTable tablaHorarios = new PdfPTable(5);
            tablaHorarios.setWidthPercentage(100);
            tablaHorarios.setWidths(new int[] { 2, 2, 2, 2, 2 });
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());

            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("HORARIO", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("ENTRADA (E)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("INICIO ALIMENTACIÓN (I/A)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("FIN ALIMENTACIÓN (F/A)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("SALIDA (S)", colorPrincipal));

            for (PlanificacionDetalleDTO d : request.getDetalle_acciones()) {
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getHorario()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getEntrada_()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getInicioComida()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getFinComida()));
                tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getSalida_()));
            }

            // TABLA DE NOMENCLATURA
            PdfPTable tablaNomenclatura = new PdfPTable(2);
            tablaNomenclatura.setWidthPercentage(100);
            tablaNomenclatura.setWidths(new int[] { 3, 7 });
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("NOMENCLATURA", colorSecundario));
            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("DESCRIPCIÓN", colorSecundario));

            for (Map<String, String> def : request.getNomenclatura()) {
                tablaNomenclatura.addCell(ReporteUtil.celdaNomenclatura(def.get("nombre"), ReporteUtil.fuenteTexto()));
                tablaNomenclatura.addCell(
                        ReporteUtil.celdaNomenclaturaDescripcion(def.get("descripcion"), ReporteUtil.fuenteTexto()));

            }

            // TABLA CONTENEDORA HORIZONTAL
            PdfPTable tablaContenedora = new PdfPTable(2);
            tablaContenedora.setWidthPercentage(100);
            tablaContenedora.setWidths(new float[] { 65, 35 });

            PdfPCell celdaIzq = new PdfPCell(tablaHorarios);
            celdaIzq.setBorder(Rectangle.NO_BORDER);
            celdaIzq.setPaddingRight(10f);

            PdfPCell celdaDer = new PdfPCell(tablaNomenclatura);
            celdaDer.setBorder(Rectangle.NO_BORDER);

            tablaContenedora.addCell(celdaIzq);
            tablaContenedora.addCell(celdaDer);
            document.add(tablaContenedora);

            document.add(Chunk.NEWLINE);

            // BLOQUES POR EMPLEADO
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
                document.add(Chunk.NEWLINE);

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
                            if (j <= 31)
                                tablaMes.addCell(ReporteUtil.celdaEncabezadoDia(j));
                            else
                                tablaMes.addCell("");
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
