package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Ciudad.CiudadDTO;
import com.casapazmino.microservicio_reportes.model.Ciudad.ReporteCiudadesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteCiudadesService {

    //METODO QUE GENERA EL REPORTE PDF
    public byte[] generarReportePDF(ReporteCiudadesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPÓ Y TAMAÑO DE LA PAGINA DEL REPORTE
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

            //TITULO DEL REPORTE
            document.add(ReporteUtil.crearTituloReporte("LISTA DE CIUDADES"));

            //COLORES DE LA EMPRESA
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            //TABLA DE LOS DATOS DE CIUDAD
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(new float[]{2, 4});
            tabla.setSpacingBefore(10f);

            //ENCABEZADOS DE LA TABLA
            tabla.addCell(ReporteUtil.crearCelda("Provincia", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("Ciudad", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            //FILAS DE LA TABLA CIUDAD
            boolean zebra = false;
            for (CiudadDTO ciudad : request.getCiudades()) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(ciudad.getProvincia(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(ciudad.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                zebra = !zebra;
            }

            document.add(tabla);
            document.close();
            writer.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
