package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Coordenada.CoordenadaDTO;
import com.casapazmino.microservicio_reportes.model.Coordenada.ReporteCoordenadasRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteCoordenadasService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteCoordenadasRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            //LOGO DE EMPRESA
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            //TITULO DE EMPRESA
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            
            //TITULO DE REPORTE
            document.add(ReporteUtil.crearTituloReporte("Lista de coordenadas geográficas"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            //TABLA
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(60);
            tabla.setWidths(new float[]{1.8f, 3f, 5.1f, 5.1f});
            tabla.setSpacingBefore(10f);

            //ENCABEZADOS DE LA TABLA
            String[] headers = {"Código", "Descripción", "Latitud", "Longitud"};
            for (String col : headers) {
                tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            }

            //FILAS DE LA TABLA
            boolean zebra = false;
            for (CoordenadaDTO c : request.getCoordenadas()) {
                Color fondo = zebra ? zebraColor : null;
                zebra = !zebra;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(c.getId()), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getDescripcion(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLatitud(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLongitud(), ReporteUtil.fuenteTablaData(), fondo));
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
