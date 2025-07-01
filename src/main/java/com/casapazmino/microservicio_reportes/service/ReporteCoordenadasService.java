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
import java.util.List;

@Service
public class ReporteCoordenadasService {

    public byte[] generarReportePDF(ReporteCoordenadasRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4, 40, 40, 50, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Título empresa y del reporte
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("Lista de coordenadas geográficas"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(100);
            tabla.setWidths(new float[]{2f, 5f, 4f, 4f});
            tabla.setSpacingBefore(5f);

            // Encabezados
            String[] headers = {"Código", "Descripción", "Latitud", "Longitud"};
            for (String col : headers) {
                tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezado(), colorPrincipal));
            }

            // Filas con zebra
            boolean zebra = false;
            for (CoordenadaDTO c : request.getCoordenadas()) {
                Color fondo = zebra ? zebraColor : null;
                zebra = !zebra;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(c.getId()), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getDescripcion(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLatitud(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(c.getLongitud(), ReporteUtil.fuenteTexto(), fondo));
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
