package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.DataHoraExtraLista;
import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.service.horasextra.interfaces.ReporteFile;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.format.DateTimeFormatter;

import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Component;

@Component("reporteHoraExtraPdf")
public class ReporteHoraExtraPdf implements ReporteFile {


    @Override
    public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request) {
        System.out.println("La data que llega es  "+ request.getData().toString());
        final float[] WIDTHS = {2f, 4f};

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(request.getUsuario(), request.getFraseMarcaAgua(), request.getColorPrincipal()));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("Horas Extra"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(40);
            tabla.setSpacingBefore(10f);
            tabla.setWidths(WIDTHS);
            tabla.setHorizontalAlignment(Element.ALIGN_CENTER);

            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Fecha", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Hora", colorPrincipal));

            boolean zebra = false;
            if (request.getData() != null && request.getData().getData() != null) {
                DateTimeFormatter fechaFormateada = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                DateTimeFormatter horaFormateada = DateTimeFormatter.ofPattern("HH:mm:ss");

                for (DataHoraExtraLista dataHoraExtra : request.getData().getData()) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;

                    String fechaString = dataHoraExtra.getFecha().format(fechaFormateada);

                    tabla.addCell(ReporteUtil.celdaDataCentro(
                            ReporteUtil.formatearFechaConDia(fechaString), fondo));

                    tabla.addCell(ReporteUtil.celdaDataCentro(
                            dataHoraExtra.getHoraEntradaReal().format(horaFormateada), fondo));

                    zebra = !zebra;
                }
            }

            document.add(tabla);

            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteHorasExtra.pdf", e);
        } finally {
            if (document != null && document.isOpen()) {
                try {
                    document.close();
                } catch (Exception ignore) {
                }
            }
            if (writer != null) {
                try {
                    writer.close();
                } catch (Exception ignore) {
                }
            }
            if (baos != null) {
                try {
                    baos.close();
                } catch (Exception ignore) {
                }
            }
        }
    }
}
