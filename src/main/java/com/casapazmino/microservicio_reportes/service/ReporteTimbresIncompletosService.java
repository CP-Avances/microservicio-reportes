package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTimbresIncompletosService {

    public byte[] generarReportePDF(ReporteTimbresIncompletosRequest request) {
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
            if (logo != null) {
                document.add(logo);
            }

            // Empresa y título
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

            String titulo = "TIMBRES INCOMPLETOS - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));

            String subtitulo = "PERIODO DEL: " + request.getPeriodo().getInicio() + " AL "
                    + request.getPeriodo().getFin();
            document.add(ReporteUtil.crearTituloPeriodo(subtitulo));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            AtomicInteger contadorGlobal = new AtomicInteger(0);

            for (TimbresSucursalDTO suc : request.getData_pdf()) {
                for (EmpleadoDTO emp : suc.getEmpleados()) {
                    if (emp.getTimbres() != null) {
                        contadorGlobal.addAndGet(emp.getTimbres().size());
                    }
                }
            }

            // Título verde - LISTA EMPLEADOS
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
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            for (TimbresSucursalDTO suc : request.getData_pdf()) {
                for (EmpleadoDTO emp : suc.getEmpleados()) {

                    // Tabla de información del empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                    infoEmpleado.addCell(
                            celdaInfoMixta("EMPLEADO:", emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado.addCell(celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);

                    PdfPCell celdaContenedora = new PdfPCell(infoEmpleado);
                    celdaContenedora.setPadding(0);
                    celdaContenedora.setBorder(Rectangle.BOX);

                    tablaContenedora.addCell(celdaContenedora);
                    document.add(tablaContenedora);

                    // Tabla de timbres
                    PdfPTable tablaTimbres = new PdfPTable(4);
                    tablaTimbres.setWidthPercentage(100);
                    tablaTimbres.setSpacingBefore(5f);
                    tablaTimbres.setWidths(new float[] { 1, 4, 4, 4 });

                    // ===== Encabezado: Fila 1 =====

                    PdfPCell celdaNum = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaNum.setRowspan(2);
                    celdaNum.setHorizontalAlignment(Element.ALIGN_CENTER);
                    celdaNum.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    celdaNum.setBorder(Rectangle.BOX);
                    tablaTimbres.addCell(celdaNum);

                    PdfPCell celdaTimbre = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal);
                    celdaTimbre.setColspan(2);
                    celdaTimbre.setHorizontalAlignment(Element.ALIGN_CENTER);
                    celdaTimbre.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    celdaTimbre.setBorder(Rectangle.TOP | Rectangle.BOTTOM);
                    tablaTimbres.addCell(celdaTimbre);

                    PdfPCell celdaAccion = ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal);
                    celdaAccion.setRowspan(2);
                    celdaAccion.setHorizontalAlignment(Element.ALIGN_CENTER);
                    celdaAccion.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    celdaAccion.setBorder(Rectangle.BOX);
                    tablaTimbres.addCell(celdaAccion);

                    // ===== Encabezado: Fila 2 =====
                    tablaTimbres
                            .addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaTimbres
                            .addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    int contadorLocal = 1;
                    for (TimbreDTO t : emp.getTimbres()) {
                        Color fondo = (contadorLocal % 2 == 0) ? zebraColor : null;

                        String[] partes = t.getFechaHora().split(" ");
                        String fecha = partes.length > 0 ? partes[0] : "";
                        String hora = partes.length > 1 ? partes[1] : "";

                        tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(fecha),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(hora, ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(traducirAccion(t.getAccion()),
                                ReporteUtil.fuenteTexto(), fondo));

                        contadorLocal++;
                        contadorGlobal.incrementAndGet();
                    }

                    document.add(tablaTimbres);
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

    private PdfPCell celdaInfoMixta(String etiqueta, String valor, Color fondo) {
        Phrase contenido = new Phrase();
        contenido.add(new Chunk(etiqueta + " ", ReporteUtil.fuenteTexto()));
        contenido.add(new Chunk(valor != null ? valor : "", ReporteUtil.fuenteTexto()));

        PdfPCell celda = new PdfPCell(contenido);
        celda.setBackgroundColor(fondo);
        celda.setPadding(5f);
        celda.setBorder(Rectangle.NO_BORDER);
        return celda;
    }

    public static String traducirAccion(String codigo) {
        if (codigo == null)
            return "";

        switch (codigo.trim().toUpperCase()) {
            case "E":
                return "Entrada";
            case "S":
                return "Salida";
            case "I/A":
                return "Inicio alimentación";
            case "F/A":
                return "Fin alimentación";
            default:
                return codigo;
        }
    }

}
