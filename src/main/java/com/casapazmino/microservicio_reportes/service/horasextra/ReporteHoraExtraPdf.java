package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.DataHoraExtra;
import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
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
    final float[] WIDTHS = { 2f, 4f };

    Document document = null;
    PdfWriter writer = null;
    ByteArrayOutputStream baos = null;

    try {
      // 1) Inicialización
      baos = new ByteArrayOutputStream();
      document = new Document(PageSize.A4);
      writer = PdfWriter.getInstance(document, baos);
      writer.setPageEvent(new ConfiguracionPaginaPDF(
          request.getUsuario(),
          request.getFraseMarcaAgua(),
          request.getColorPrincipal()));
      document.open();

      // 2) Construcción
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

      // Encabezados
      tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Fecha", colorPrincipal));
      tabla.addCell(ReporteUtil.celdaEncabezadoTabla("Hora", colorPrincipal));

      // Cuerpo (zebra)
      boolean zebra = false;
      if (request.getDataHoraExtra() != null) {
        for (DataHoraExtra dataHoraExtra : request.getDataHoraExtra()) {
          Color fondo = zebra ? colorZebra : Color.WHITE;
          tabla.addCell(
              ReporteUtil.celdaDataCentro(ReporteUtil.formatearFechaConDia(dataHoraExtra.getFecha()), fondo));
          tabla.addCell(ReporteUtil.celdaDataCentro(dataHoraExtra.getHora(), fondo));
          zebra = !zebra;
        }
      }

      document.add(tabla);

      // 3) Cierre y retorno
      document.close();
      return baos.toByteArray();

    } catch (IllegalArgumentException e) {
      // Si algún helper valida y falla, el controller puede mapearlo a 400
      throw e;
    } catch (Exception e) {
      // 500 interno uniforme
      throw new ReportBuildException("No se pudo generar ReporteHorasExtra.pdf", e);
    } finally {
      // 4) Ciclo de recursos garantizado
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
