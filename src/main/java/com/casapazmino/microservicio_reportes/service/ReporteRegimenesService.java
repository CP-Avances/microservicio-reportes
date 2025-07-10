package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.RegimenLaboral.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteRegimenesService {

    public byte[] generarReporteRegimenesPDF(ReporteRegimenesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate());
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
            document.add(ReporteUtil.crearTituloReporte("RÉGIMEN LABORAL"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebra = ReporteUtil.colorZebraClaro();

            for (RegimenDTO reg : request.getRegimenes()) {
                int contador = 0;
                PdfPTable tablaCabecera = new PdfPTable(4);
                tablaCabecera.setWidthPercentage(100);
                tablaCabecera.setWidths(new float[] { 2.5f, 3, 2.5f, 2 });

                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("PAÍS: ", reg.getPais(), colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN: ", reg.getDescripcion(), colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("CONTINUIDAD LABORAL: ",(reg.getContinuidad_laboral() ? "SÍ" : "NO"), colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("CÓDIGO: ", String.valueOf(reg.getId()), colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("PERIODO LABORAL: ", reg.getMes_periodo() + " Meses",colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("DÍAS POR MES: ", String.valueOf(reg.getDias_mes()),colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("TIEMPO MÍNIMO: ",
                    (reg.getTrabajo_minimo_mes() > 0 ? reg.getTrabajo_minimo_mes() + " Meses": reg.getTrabajo_minimo_horas() + " Horas"),colorPrincipal));
                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("ANTIGÜEDAD LABORAL: ",(reg.getAntiguedad() ? "SÍ" : "NO"), colorPrincipal));

                PdfPTable tablaContenedora = new PdfPTable(1);
                tablaContenedora.setWidthPercentage(100);
                PdfPCell contenedor1 = new PdfPCell(tablaCabecera);
                contenedor1.setPadding(0);
                contenedor1.setBorder(Rectangle.BOX);
                tablaContenedora.addCell(contenedor1);
                tablaContenedora.setSpacingAfter(6f);
                tablaContenedora.setSpacingBefore(40f);
                document.add(tablaContenedora);

                // Tabla 1 - Configuración de Vacaciones
                PdfPTable configVac = new PdfPTable(2);
                configVac.setWidthPercentage(100);
                configVac.setWidths(new float[] { 2.5f, 1.5f });
                configVac.addCell(ReporteUtil.crearCelda("CONFIGURACIÓN DE VACACIONES", ReporteUtil.fuenteEncabezado(),colorPrincipal, 1, 2));
                configVac.addCell(ReporteUtil.celdaCentro("DÍAS HÁBILES", colorSecundario));
                configVac.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_laboral()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                configVac.addCell(ReporteUtil.celdaCentro("DÍAS LIBRES", colorSecundario));
                configVac.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_libre()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                configVac.addCell(ReporteUtil.celdaCentro("DÍAS CALENDARIO", colorSecundario));
                configVac.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_calendario()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                configVac.addCell(ReporteUtil.celdaCentro("ACUMULA VACACIONES", colorSecundario));
                configVac.addCell(ReporteUtil.celdaCentro(reg.getAcumular() ? "SÍ" : "NO",(contador++ % 2 == 0) ? zebra : Color.WHITE));
                if (reg.getAcumular()) {
                    configVac.addCell(ReporteUtil.celdaCentro("MÁXIMO DÍAS ACUMULABLES", colorSecundario));
                    configVac.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getDias_maximo_acumulacion()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                }
                configVac.addCell(ReporteUtil.celdaCentro("VACACIONES POR PERÍODOS", colorSecundario));
                configVac.addCell(ReporteUtil.celdaCentro(reg.getVacacion_divisible() ? "SÍ" : "NO",(contador++ % 2 == 0) ? zebra : Color.WHITE));
                if (reg.getVacacion_divisible() && reg.getPeriodos_vacacionales() != null
                        && !reg.getPeriodos_vacacionales().isEmpty()) {
                    for (PeriodoVacacionalDTO p : reg.getPeriodos_vacacionales()) {
                        configVac.addCell(ReporteUtil.celdaCentro(p.getDescripcion(),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                        configVac.addCell(ReporteUtil.celdaCentro(p.getDias_vacacion() + " días",(contador++ % 2 == 0) ? zebra : Color.WHITE));
                    }
                }

                // Tabla 2 - Vacaciones Ganadas
                PdfPTable vacGanadas = new PdfPTable(2);
                vacGanadas.setWidthPercentage(100);
                vacGanadas.setWidths(new float[] { 2.5f, 1.5f });
                vacGanadas.addCell(ReporteUtil.crearCelda("VACACIONES GANADAS", ReporteUtil.fuenteEncabezado(),colorPrincipal, 1, 2));
                vacGanadas.addCell(ReporteUtil.celdaCentro("POR MES (HÁBILES)", colorSecundario));
                vacGanadas.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_laboral_mes()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                vacGanadas.addCell(ReporteUtil.celdaCentro("POR MES (CALENDARIO)", colorSecundario));
                vacGanadas.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_calendario_mes()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                vacGanadas.addCell(ReporteUtil.celdaCentro("POR DÍA (HÁBILES)", colorSecundario));
                vacGanadas.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getLaboral_dias()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                vacGanadas.addCell(ReporteUtil.celdaCentro("POR DÍA (CALENDARIO)", colorSecundario));
                vacGanadas.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getCalendario_dias()),(contador++ % 2 == 0) ? zebra : Color.WHITE));

                // Tabla 3 - Antigüedad
                PdfPTable antiguedad = new PdfPTable(2);
                antiguedad.setWidthPercentage(100);
                antiguedad.setWidths(new float[] { 2.5f, 1.5f });
                antiguedad.addCell(ReporteUtil.crearCelda("CONFIGURACIÓN DE ANTIGÜEDAD", ReporteUtil.fuenteEncabezado(),colorPrincipal, 1, 2));
                if (Boolean.TRUE.equals(reg.getAntiguedad())) {
                    if (Boolean.TRUE.equals(reg.getAntiguedad_fija())) {
                        antiguedad.addCell(ReporteUtil.celdaCentro("TIPO", colorSecundario));
                        antiguedad.addCell(ReporteUtil.celdaCentro("FIJA", (contador++ % 2 == 0) ? zebra : Color.WHITE));
                        antiguedad.addCell(ReporteUtil.celdaCentro("AÑOS ANTIGÜEDAD", colorSecundario));
                        antiguedad.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getAnio_antiguedad()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                        antiguedad.addCell(ReporteUtil.celdaCentro("DÍAS ADICIONALES", colorSecundario));
                        antiguedad.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getDias_antiguedad()),(contador++ % 2 == 0) ? zebra : Color.WHITE));
                    } else if (Boolean.TRUE.equals(reg.getAntiguedad_variable()) && reg.getRangos_antiguedad() != null&& !reg.getRangos_antiguedad().isEmpty()) {
                        antiguedad.addCell(ReporteUtil.celdaCentro("TIPO", colorSecundario));
                        antiguedad.addCell(ReporteUtil.celdaCentro("VARIABLE", (contador++ % 2 == 0) ? zebra : Color.WHITE));
                        for (RangoAntiguedadDTO r : reg.getRangos_antiguedad()) {
                            antiguedad.addCell(ReporteUtil.celdaCentro("Desde " + r.getAnio_desde() + " hasta " + r.getAnio_hasta() + " años",colorSecundario));
                            antiguedad.addCell(ReporteUtil.celdaCentro(r.getDias_antiguedad() + " días",(contador++ % 2 == 0) ? zebra : Color.WHITE));
                        }
                    }
                } else {
                    PdfPCell celdaNoAplica = ReporteUtil.celdaCentro("NO APLICA",(contador++ % 2 == 0) ? zebra : Color.WHITE);
                    celdaNoAplica.setColspan(2);
                    antiguedad.addCell(celdaNoAplica);

                }

                // Medir número de filas por tabla
                int filasVac = configVac.size();
                int filasGanadas = vacGanadas.size();
                int filasAntig = antiguedad.size();
                int maxFilas = Math.max(filasVac, Math.max(filasGanadas, filasAntig));

                // Rellenar dinámicamente con filas vacías sin bordes visibles
                while (configVac.size() < maxFilas) {
                    PdfPCell celda1 = ReporteUtil.celdaCentro("", Color.WHITE);
                    celda1.setBorder(Rectangle.NO_BORDER);
                    PdfPCell celda2 = ReporteUtil.celdaCentro("", Color.WHITE);
                    celda2.setBorder(Rectangle.NO_BORDER);
                    configVac.addCell(celda1);
                    configVac.addCell(celda2);
                }
                while (vacGanadas.size() < maxFilas) {
                    PdfPCell celda1 = ReporteUtil.celdaCentro("", Color.WHITE);
                    celda1.setBorder(Rectangle.NO_BORDER);
                    PdfPCell celda2 = ReporteUtil.celdaCentro("", Color.WHITE);
                    celda2.setBorder(Rectangle.NO_BORDER);
                    vacGanadas.addCell(celda1);
                    vacGanadas.addCell(celda2);
                }
                while (antiguedad.size() < maxFilas) {
                    PdfPCell celda1 = ReporteUtil.celdaCentro("", Color.WHITE);
                    celda1.setBorder(Rectangle.NO_BORDER);
                    PdfPCell celda2 = ReporteUtil.celdaCentro("", Color.WHITE);
                    celda2.setBorder(Rectangle.NO_BORDER);
                    antiguedad.addCell(celda1);
                    antiguedad.addCell(celda2);
                }

                PdfPCell celdaConfigVac = new PdfPCell(configVac);
                celdaConfigVac.setVerticalAlignment(Element.ALIGN_TOP);
                celdaConfigVac.setPadding(0);
                celdaConfigVac.setBorder(Rectangle.NO_BORDER);

                PdfPCell celdaVacGanadas = new PdfPCell(vacGanadas);
                celdaVacGanadas.setVerticalAlignment(Element.ALIGN_TOP);
                celdaVacGanadas.setPadding(0);
                celdaVacGanadas.setBorder(Rectangle.NO_BORDER);

                PdfPCell celdaAntiguedad = new PdfPCell(antiguedad);
                celdaAntiguedad.setVerticalAlignment(Element.ALIGN_TOP);
                celdaAntiguedad.setPadding(0);
                celdaAntiguedad.setBorder(Rectangle.NO_BORDER);

                PdfPTable contenedor = new PdfPTable(3);
                contenedor.setWidthPercentage(100);
                contenedor.setWidths(new float[] { 3f, 3f, 3f });
                contenedor.addCell(celdaConfigVac);
                contenedor.addCell(celdaVacGanadas);
                contenedor.addCell(celdaAntiguedad);
                contenedor.setSpacingAfter(10f);

                contador++;
                document.add(contenedor);
            }
            document.close();
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}