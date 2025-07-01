package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteFaltas.EmpleadoFaltasDTO;
import com.casapazmino.microservicio_reportes.model.ReporteFaltas.FaltaDTO;
import com.casapazmino.microservicio_reportes.model.ReporteFaltas.GrupoFaltasDTO;
import com.casapazmino.microservicio_reportes.model.ReporteFaltas.ReporteFaltasRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteFaltasService {

    public byte[] generarReporteFaltasPDF(ReporteFaltasRequest request) {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try {
            Document document = new Document(PageSize.A4, 40, 40, 50, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(
                    "FALTAS - USUARIOS " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS")));
            document.add(ReporteUtil
                    .crearTituloReporte("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));
            document.add(Chunk.NEWLINE);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            int totalFaltasGeneral = 0;

            for (GrupoFaltasDTO grupo : request.getGrupos()) {
                // Cabecera
                PdfPTable tablaCabecera = new PdfPTable(3);
                tablaCabecera.setWidthPercentage(100);
                tablaCabecera.setWidths(new float[] { 4, 4, 2 });
                tablaCabecera.setSpacingBefore(10f);

                String descripcion = obtenerDescripcionFiltro(request, grupo);
                String establecimiento = request.isResumen() ? "SUCURSAL: " + grupo.getSucursal() : "";
                int totalGrupo = grupo.getEmpleados().stream().mapToInt(e -> e.getFaltas().size()).sum();

                tablaCabecera.addCell(ReporteUtil.crearCelda(descripcion, ReporteUtil.fuenteTexto(), colorSecundario));
                tablaCabecera
                        .addCell(ReporteUtil.crearCelda(establecimiento, ReporteUtil.fuenteTexto(), colorSecundario));
                tablaCabecera.addCell(ReporteUtil.celdaEncabezado("N° Registros: " + totalGrupo, colorSecundario));

                document.add(tablaCabecera);

                // Empleados
                for (EmpleadoFaltasDTO emp : grupo.getEmpleados()) {
                    // Datos personales
                    PdfPTable tablaEmpleado = new PdfPTable(3);
                    tablaEmpleado.setWidthPercentage(100);
                    tablaEmpleado.setWidths(new float[] { 4, 6, 3 });
                    tablaEmpleado.setSpacingBefore(6f);

                    tablaEmpleado.addCell(ReporteUtil.crearCelda("C.C.: " + emp.getIdentificacion(),
                            ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado
                            .addCell(ReporteUtil.crearCelda("EMPLEADO: " + emp.getApellido() + " " + emp.getNombre(),
                                    ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("DEPARTAMENTO: " + emp.getDepartamento(),
                            ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("CORREO: " + emp.getCorreo(),
                            ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("GÉNERO: " + emp.getGenero(),
                            ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(
                            ReporteUtil.crearCelda("CARGO: " + emp.getCargo(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(ReporteUtil.crearCelda("RÉGIMEN: " + emp.getRegimen(),
                            ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(
                            ReporteUtil.crearCelda("COD: " + emp.getCodigo(), ReporteUtil.fuenteTexto(), colorZebra));
                    tablaEmpleado.addCell(
                            ReporteUtil.crearCelda("ROL: " + emp.getRol(), ReporteUtil.fuenteTexto(), colorZebra));

                    document.add(tablaEmpleado);

                    // Tabla de faltas
                    PdfPTable tablaFaltas = new PdfPTable(2);
                    tablaFaltas.setWidthPercentage(100);
                    tablaFaltas.setWidths(new float[] { 1, 5 });
                    tablaFaltas.setSpacingBefore(4f);

                    tablaFaltas.addCell(ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaFaltas
                            .addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    int count = 0;
                    for (FaltaDTO falta : emp.getFaltas()) {
                        count++;
                        Color bg = (count % 2 == 0) ? colorZebra : null;
                        tablaFaltas
                                .addCell(ReporteUtil.crearCelda(String.valueOf(count), ReporteUtil.fuenteTexto(), bg));
                        tablaFaltas.addCell(ReporteUtil.crearCelda(falta.getFecha(), ReporteUtil.fuenteTexto(), bg));
                    }

                    tablaFaltas
                            .addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteEncabezado(), colorSecundario));
                    tablaFaltas.addCell(ReporteUtil.crearCelda(String.valueOf(count), ReporteUtil.fuenteEncabezado(),
                            colorSecundario));

                    totalFaltasGeneral += count;
                    document.add(tablaFaltas);
                }
            }

            // Resumen final
            if (!request.isResumen()) {
                document.add(Chunk.NEWLINE);
                PdfPTable resumen = new PdfPTable(3);
                resumen.setWidthPercentage(100);
                resumen.setWidths(new float[] { 4, 4, 2 });
                resumen.setSpacingBefore(10f);

                resumen.addCell(
                        ReporteUtil.crearCelda("TOTAL GENERAL", ReporteUtil.fuenteEncabezado(), colorSecundario));
                resumen.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorSecundario));
                resumen.addCell(ReporteUtil.crearCelda(String.valueOf(totalFaltasGeneral),
                        ReporteUtil.fuenteEncabezado(), colorSecundario));

                document.add(resumen);
            }

            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return baos.toByteArray();
    }

    private String obtenerDescripcionFiltro(ReporteFaltasRequest request, GrupoFaltasDTO grupo) {
        if (request.isResumen())
            return "LISTA EMPLEADOS";
        switch (request.getOpcionBusqueda()) {
            case 1:
                return "RÉGIMEN LABORAL: " + grupo.getNombre();
            case 2:
                return "DEPARTAMENTO: " + grupo.getDepartamento();
            case 3:
                return "CARGO: " + grupo.getNombre();
            case 4:
                return "CIUDAD: " + grupo.getCiudad();
            default:
                return "";
        }
    }
}
