package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTimbresVirtualesService {

    public byte[] generarReportePDF(ReporteTimbresVirtualesRequest request) {
        try {
            Rectangle orientacion = request.isTimbreDispositivo() ? PageSize.A4.rotate() : PageSize.A4;

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(orientacion, 40, 40, 30, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            // Logo y encabezados
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            String titulo = "TIMBRES VIRTUALES - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil.crearTituloPeriodo("PERIODO DEL: " + request.getPeriodo().getInicio() + " AL " + request.getPeriodo().getFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getData_pdf().forEach(grupo ->
                    grupo.getEmpleados().forEach(emp ->
                            contadorGlobal.addAndGet(emp.getTimbres().size())));

            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(new float[]{8, 2});
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                for (EmpleadoTimbreDTO emp : grupo.getEmpleados()) {

                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[]{4, 4, 4});

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    document.add(tablaContenedora);

                    PdfPTable tablaTimbres = new PdfPTable(8);
                    tablaTimbres.setWidthPercentage(100);
                    tablaTimbres.setSpacingBefore(5f);
                    tablaTimbres.setWidths(new float[]{1, 2.5f, 1.5f, 1.5f, 2.3f, 4.5f, 2, 2});

                    // Encabezado Fila 1
                    PdfPCell celdaN = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaN.setRowspan(2); celdaN.setHorizontalAlignment(Element.ALIGN_CENTER);
                    tablaTimbres.addCell(celdaN);

                    PdfPCell celdaTimbre = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaTimbre.setColspan(2); celdaTimbre.setHorizontalAlignment(Element.ALIGN_CENTER);
                    tablaTimbres.addCell(celdaTimbre);

                    PdfPCell celdaReloj = ReporteUtil.crearCelda("RELOJ", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaReloj.setRowspan(2); tablaTimbres.addCell(celdaReloj);

                    PdfPCell celdaAccion = ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaAccion.setRowspan(2); tablaTimbres.addCell(celdaAccion);

                    PdfPCell celdaObs = ReporteUtil.crearCelda("OBSERVACIÓN", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaObs.setRowspan(2); tablaTimbres.addCell(celdaObs);

                    PdfPCell celdaLong = ReporteUtil.crearCelda("LONGITUD", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaLong.setRowspan(2); tablaTimbres.addCell(celdaLong);

                    PdfPCell celdaLat = ReporteUtil.crearCelda("LATITUD", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaLat.setRowspan(2); tablaTimbres.addCell(celdaLat);

                    // Encabezado Fila 2
                    tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaTimbres.addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    // Cuerpo
                    int contadorLocal = 1;
                    for (TimbreUsuarioDTO t : emp.getTimbres()) {
                        Color fondo = (contadorLocal % 2 == 0) ? zebraColor : Color.WHITE;

                        String[] partes = t.getFecha_hora_timbre().split(" ");
                        String fecha = partes.length > 0 ? partes[0] : "";
                        String hora = partes.length > 1 ? partes[1] : "";

                        tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(fecha), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(hora, ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getId_reloj(), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.traducirAccion(t.getAccion()), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getObservacion(), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getLongitud(), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getLatitud(), ReporteUtil.fuenteTexto(), fondo));

                        contadorLocal++;
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
}
