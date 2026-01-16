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


        final float[] WIDTHS = {2.5f, 2f, 2f, 2f, 2f, 3f};

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            // ================= LOGO =================
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // ================= TÍTULOS =================
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("Reporte de Horas Extra"));

            // ================= COLORES =================
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // ================= TABLA =================
            PdfPTable tabla = new PdfPTable(6);
            tabla.setWidthPercentage(100);
            tabla.setSpacingBefore(10f);
            tabla.setWidths(WIDTHS);
            tabla.setHorizontalAlignment(Element.ALIGN_CENTER);

            // ---------- ENCABEZADOS ----------
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Fecha", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Entrada", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Salida", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Min. Extra", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Horas Extra", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Estado", colorPrincipal));

            // ---------- FORMATEADORES ----------
            DateTimeFormatter fechaFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
            DateTimeFormatter horaFormatter = DateTimeFormatter.ofPattern("HH:mm");

            boolean zebra = false;

            if (request.getData() != null && request.getData().getData() != null) {
                for (DataHoraExtraLista dataHoraExtra : request.getData().getData()) {

                    Color fondo = zebra ? colorZebra : Color.WHITE;

                    String fecha = ReporteUtil.formatearFechaConDia(
                            dataHoraExtra.getFecha().format(fechaFormatter));

                    String horaEntrada = dataHoraExtra.getHoraEntradaReal() != null
                            ? dataHoraExtra.getHoraEntradaReal().format(horaFormatter)
                            : "--";

                    String horaSalida = dataHoraExtra.getHoraSalidaReal() != null
                            ? dataHoraExtra.getHoraSalidaReal().format(horaFormatter)
                            : "--";

                    String minutosExtra = dataHoraExtra.getMinutosHorasExtra() != null
                            ? String.valueOf(dataHoraExtra.getMinutosHorasExtra())
                            : "0";

                    String horasExtra = dataHoraExtra.getHorasHorasExtra() != null
                            ? dataHoraExtra.getCodigoEmpleado()
                            : "0.00";

                    String estado = dataHoraExtra.getEstadoCalculoDesc() != null
                            ? dataHoraExtra.getEstadoCalculoDesc()
                            : "--";

                    tabla.addCell(ReporteUtil.celdaDataCentro(fecha, fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(horaEntrada, fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(horaSalida, fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(minutosExtra, fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(horasExtra, fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(estado, fondo));

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
