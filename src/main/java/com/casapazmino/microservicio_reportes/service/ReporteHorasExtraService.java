package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Genero.GeneroDTO;
import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class ReporteHorasExtraService {
    public byte[] generarReporteGenerosPDF(ReporteHorasExtraRequest request) {
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
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("TEST", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("TEST", colorPrincipal));

            // Cuerpo (zebra)
            List<GeneroDTO> generos =  new ArrayList<>();
            boolean zebra = false;
            if (generos != null) {
                for (GeneroDTO g : generos) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(
                            ReporteUtil.celdaDataCentro(String.valueOf(g.getId()), fondo));
                    tabla.addCell(ReporteUtil.celdaDataCentro(g.getGenero(), fondo));
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
            throw new ReportBuildException("No se pudo generar ReporteGeneros.pdf", e);
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
