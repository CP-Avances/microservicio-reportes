package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ResumenAsistencia.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteAsistenciaService {

    public byte[] generarReportePDF(ReporteAsistenciaRequest request) {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try {
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 50, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();
            System.out.println("→ Abriendo documento PDF...");

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

            String titulo = "RESUMEN DE ASISTENCIA - " +
                    (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));

            document.add(ReporteUtil.crearTituloReporte("PERIODO DEL: " +
                    request.getFechaInicio() + " AL " + request.getFechaFin()));
            document.add(Chunk.NEWLINE);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            List<GrupoAsistenciaDTO> grupos = request.getGrupos();
            System.out.println("→ Grupos recibidos: " + grupos.size());

            for (GrupoAsistenciaDTO grupo : grupos) {
                System.out.println("→ Grupo: " + grupo.getNombre());

                PdfPTable tablaCabecera = new PdfPTable(3);
                tablaCabecera.setWidthPercentage(100);
                tablaCabecera.setSpacingBefore(5f);
                tablaCabecera.setWidths(new float[]{4, 4, 2});

                tablaCabecera.addCell(ReporteUtil.crearCelda("AGRUPADO POR: " + grupo.getNombre(), ReporteUtil.fuenteTexto(), null));
                tablaCabecera.addCell(ReporteUtil.crearCelda("SUCURSAL: " + grupo.getSucursal(), ReporteUtil.fuenteTexto(), null));
                tablaCabecera.addCell(ReporteUtil.crearCelda("N° Empleados: " + grupo.getEmpleados().size(), ReporteUtil.fuenteTexto(), null));

                document.add(tablaCabecera);
                document.add(Chunk.NEWLINE);

                for (EmpleadoAsistenciaDTO emp : grupo.getEmpleados()) {
                    System.out.println("   → Empleado: " + emp.getNombre() + " " + emp.getApellido());

                    PdfPTable tablaEmpleado = new PdfPTable(3);
                    tablaEmpleado.setWidthPercentage(100);
                    tablaEmpleado.setSpacingBefore(2f);
                    tablaEmpleado.setWidths(new float[]{4, 6, 3});
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("C.C.: " + emp.getIdentificacion(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("EMPLEADO: " + emp.getApellido() + " " + emp.getNombre(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("COD: " + emp.getCodigo(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("RÉGIMEN: " + emp.getRegimen(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("DEPARTAMENTO: " + emp.getDepartamento(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("CARGO: " + emp.getCargo(), ReporteUtil.fuenteTexto(), colorZebra));

                    document.add(tablaEmpleado);

                    List<RegistroAsistenciaDTO> registros = emp.getTLaborado();
                    System.out.println("      ⚙️  Registros recibidos desde DTO: " + registros);

                    if (registros != null && !registros.isEmpty()) {
                        System.out.println("      ✅ Cantidad de registros: " + registros.size());

                        PdfPTable tablaDetalle = new PdfPTable(8);
                        tablaDetalle.setWidthPercentage(100);
                        tablaDetalle.setSpacingBefore(4f);
                        tablaDetalle.setWidths(new float[]{3, 3, 3, 3, 3, 3, 3, 6});

                        String[] headers = {"FECHA", "ENTRADA", "SALIDA", "ATRASO", "SALIDA ANTIC.", "T. ALIM. ASIG.", "T. ALIM. TOMADO", "TIEMPO LABORADO"};
                        for (String h : headers) {
                            tablaDetalle.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), colorPrincipal));
                        }

                        for (int i = 0; i < registros.size(); i++) {
                            RegistroAsistenciaDTO r = registros.get(i);
                            Color bg = (i % 2 == 0) ? colorZebra : null;

                            String fecha = safe(r.getEntrada() != null ? r.getEntrada().getFecha_horario() : "");
                            String entrada = safe(r.getEntrada() != null ? r.getEntrada().getFecha_hora_timbre() : "");
                            String salida = safe(r.getSalida() != null ? r.getSalida().getFecha_hora_timbre() : "");

                            String atraso = convertirMinutosAHHMM(r.getMinAtrasos());
                            String salidaAnt = convertirMinutosAHHMM(r.getMinSalidasAnticipadas());
                            String alimAsig = r.getInicioAlimentacion() != null ? convertirMinutosAHHMM(r.getInicioAlimentacion().getMinutos_alimentacion()) : "00:00";
                            String alimTomado = convertirMinutosAHHMM(r.getMinAlimentacion());
                            String laborado = convertirMinutosAHHMM(r.isControl() ? r.getMinLaborados() : r.getMinPlanificados());

                            tablaDetalle.addCell(ReporteUtil.crearCelda(fecha, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(entrada, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(salida, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(atraso, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(salidaAnt, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(alimAsig, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(alimTomado, ReporteUtil.fuenteTexto(), bg));
                            tablaDetalle.addCell(ReporteUtil.crearCelda(laborado, ReporteUtil.fuenteTexto(), bg));
                        }

                        document.add(tablaDetalle);
                        document.add(Chunk.NEWLINE);
                    } else {
                        System.out.println("   ⚠️  Sin registros para este empleado.");
                    }
                }
            }

            document.close();
            System.out.println("→ Documento cerrado correctamente.");

        } catch (Exception e) {
            System.out.println("⚠️ ERROR al generar el PDF:");
            e.printStackTrace();
        }

        System.out.println("✅ Tamaño del PDF generado (bytes): " + baos.toByteArray().length);
        return baos.toByteArray();
    }

    private String convertirMinutosAHHMM(int minutos) {
        int horas = minutos / 60;
        int min = minutos % 60;
        return String.format("%02d:%02d", horas, min);
    }

    private String safe(String texto) {
        return texto != null ? texto : "";
    }
}
