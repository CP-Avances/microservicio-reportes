package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Parametro.DetalleParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteParametrosService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReporteParametrosPDF(ReporteParametrosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            //LOGO
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            //TITULO DE EMPRESA (EJM. CASA PAZMIÑO S.A.)
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

            //TITULO DE REPORTE (EJM. REPORTE DE ATRASOS)
            document.add(ReporteUtil.crearTituloReporte("PARÁMETROS GENERALES"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

            for (ParametroDTO parametro : request.getParametros()) {
                // TABLA DE ENCABEZADOS DE LOS PARAMETROS
                PdfPTable encabezado = new PdfPTable(2);
                encabezado.setWidthPercentage(100);
                encabezado.setWidths(new float[] { 6.9f, 2.1f });

                PdfPCell celdaParametro = new PdfPCell(
                        new Phrase("PARÁMETRO: " + parametro.getDescripcion(), ReporteUtil.fuenteEncabezado()));
                celdaParametro.setBackgroundColor(colorPrincipal);
                celdaParametro.setPadding(2f);
                celdaParametro.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                encabezado.addCell(celdaParametro);

                PdfPCell celdaCodigo = new PdfPCell(
                        new Phrase("CÓDIGO PARAMETRO: " + parametro.getId(), ReporteUtil.fuenteEncabezado()));
                celdaCodigo.setBackgroundColor(colorPrincipal);
                celdaCodigo.setPadding(2f);
                celdaCodigo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                encabezado.addCell(celdaCodigo);

                encabezado.setSpacingAfter(3f);
                encabezado.setSpacingBefore(12.5f);
                document.add(encabezado);

                if (parametro.getDetalles() != null && !parametro.getDetalles().isEmpty()) {
                    //TABLA DE INFORMACION DE PARAMETROS
                    PdfPTable tabla = new PdfPTable(3);
                    tabla.setWidthPercentage(100);
                    tabla.setWidths(new float[] { 2f, 3f, 6f });

                    //ENCABEZADOS
                    String[] headers = { "CÓDIGO DETALLE", "DETALLE", "DESCRIPCIÓN" };
                    for (String h : headers) {
                        tabla.addCell(
                                ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                    }
                    //DETALLES DE PARAMETROS
                    for (DetalleParametroDTO d : parametro.getDetalles()) {
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()), ReporteUtil.fuenteTablaData(),Color.WHITE));
                        tabla.addCell(ReporteUtil.crearCelda(d.getDescripcion(), ReporteUtil.fuenteTablaData(), Color.WHITE));
                        tabla.addCell(ReporteUtil.crearCelda(d.getObservacion(), ReporteUtil.fuenteTablaData(), Color.WHITE));
                    }
                    document.add(tabla);
                }
            }
            document.close();
            writer.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
