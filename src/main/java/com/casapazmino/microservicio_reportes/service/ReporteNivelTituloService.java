package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.NivelTitulo.NivelTituloDTO;
import com.casapazmino.microservicio_reportes.model.NivelTitulo.ReporteNivelesTitulosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteNivelTituloService {

    //METODO QUE GENER EL PDF
    public byte[] generarReporteNivelTituloPDF(ReporteNivelesTitulosRequest request) {
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
            
            //TITULO DEL REPORTE
            document.add(ReporteUtil.crearTituloReporte("LISTA DE NIVELES DE TÍTULOS PROFESIONALES"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            //TABLA
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(40);
            tabla.setWidths(new float[]{2, 6});
            tabla.setSpacingBefore(10f);

            //ENCABEZADOS DE LA TABLA
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            //FILAS DE LA TABLA (CUERPO)
            List<NivelTituloDTO> lista = request.getNivelesTitulos();
            boolean zebra= false;
            for (NivelTituloDTO n : lista) {
                Color bgColor = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(n.getId()), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(n.getNombre(), ReporteUtil.fuenteTablaData(), bgColor));
                zebra = !zebra;
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
