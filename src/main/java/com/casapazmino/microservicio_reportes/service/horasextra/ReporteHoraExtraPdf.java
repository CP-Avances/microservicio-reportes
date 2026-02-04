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
import java.util.List;
import java.util.Optional;

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
            PdfPTable tabla = new PdfPTable(8);
            tabla.setWidthPercentage(100);
            tabla.setSpacingBefore(10f);
            tabla.setWidths(WIDTHS);
            tabla.setHorizontalAlignment(Element.ALIGN_CENTER);

            // ---------- ENCABEZADOS ----------
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Nombre",colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Tipo Hora Extra",colorPrincipal));
            //ESPEACIO EN BLANCO
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("", colorPrincipal));
            //ESPEACIO EN BLANCO
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("", colorPrincipal));
            //ESPEACIO EN BLANCO
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("", colorPrincipal));

            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Dia",colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Timbre Entrada",colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Timbre Salida",colorPrincipal));

            // ---------- FORMATEADORES ----------
            DateTimeFormatter fechaFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
            DateTimeFormatter horaFormatter = DateTimeFormatter.ofPattern("HH:mm");


            Optional<List<DataHoraExtraLista>> optionalDataHoraExtra = Optional.of(request.getData().getData());
            if(optionalDataHoraExtra.isEmpty()){
                throw new ReportBuildException("No se pudo generar ReporteHorasExtra.pdf");
            }
            this.generarReporteHorasExtraPdf(document, tabla, request, colorPrincipal, colorZebra, fechaFormatter, horaFormatter);

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

    private void generarReporteHorasExtraPdf(
            Document document,
            PdfPTable tabla,
            ReporteHorasExtraRequest request,
            Color colorPrincipal,
            Color colorZebra,
            DateTimeFormatter fechaFormatter,
            DateTimeFormatter horaFormatter
    ) {

        List<DataHoraExtraLista> dataList = request.getData().getData();

        if (dataList == null || dataList.isEmpty()) {
            throw new ReportBuildException("No existen datos para generar el reporte de horas extra");
        }

        boolean zebra = false;
        int size = dataList.size();

        for (int index = 0; index < size; index++) {

            DataHoraExtraLista dataHoraExtra = dataList.get(index);
            boolean esUltimo = (index == size - 1);
            Color fondo = zebra ? colorZebra : Color.WHITE;

            String fecha = dataHoraExtra.getFecha() != null
                    ? ReporteUtil.formatearFechaConDia(dataHoraExtra.getFecha().format(fechaFormatter))
                    : "--";

            String horaEntrada = dataHoraExtra.getHoraEntradaReal() != null
                    ? dataHoraExtra.getHoraEntradaReal().format(horaFormatter)
                    : "--";

            String horaSalida = dataHoraExtra.getHoraSalidaReal() != null
                    ? dataHoraExtra.getHoraSalidaReal().format(horaFormatter)
                    : "--";

            String horasExtra = dataHoraExtra.getHorasHorasExtra() != null
                    ? dataHoraExtra.getHorasHorasExtra().toString()
                    : "0.00";

            String estado = Optional.ofNullable(dataHoraExtra.getEstadoCalculoDesc()).orElse("--");
            String nombreEmpleado = Optional.ofNullable(dataHoraExtra.getNombreCompleto()).orElse("--");
            String dia = Optional.ofNullable(dataHoraExtra.getDia()).orElse("--");

            // ========= CELDAS =========
            tabla.addCell(ReporteUtil.celdaDataCentro(nombreEmpleado, colorPrincipal  ));
            tabla.addCell(ReporteUtil.celdaDataCentro("Hora Extra", colorPrincipal)); // Tipo Hora Extra (si aplica lógica luego)
            tabla.addCell(ReporteUtil.celdaDataCentro("", colorPrincipal)); // espacio
            tabla.addCell(ReporteUtil.celdaDataCentro("", colorPrincipal)); // espacio
            tabla.addCell(ReporteUtil.celdaDataCentro("", colorPrincipal)); // espacio
            tabla.addCell(ReporteUtil.celdaDataCentro(dia, colorPrincipal));
            tabla.addCell(ReporteUtil.celdaDataCentro(horaEntrada, colorPrincipal));
            tabla.addCell(ReporteUtil.celdaDataCentro(horaSalida, colorPrincipal));

            zebra = !zebra;
        }

        try {
            document.add(tabla);
        } catch (Exception e) {
            throw new ReportBuildException("Error al agregar tabla al documento PDF", e);
        }
    }

}
