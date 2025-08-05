package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Departamento.DepartamentoDTO;
import com.casapazmino.microservicio_reportes.model.Departamento.ReporteDepartamentosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteDepartamentosService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteDepartamentosRequest request) {
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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DEPARTAMENTOS"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            //TABLA DE LOS DATOS DE DEPARTAMENTOS
            PdfPTable tabla = new PdfPTable(5);
            tabla.setWidthPercentage(90);
            tabla.setWidths(new float[]{1.5f, 5, 4, 1.5f, 4});
            tabla.setSpacingBefore(10f);

            //ENCABEZADOS DE LA TABLA
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("SUCURSAL/ ESTABLECIMIENTO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO SUPERIOR", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            //FILAS CON DATOS DE LA TABLA (CUERPO)
            List<DepartamentoDTO> lista = request.getDepartamentos();
            boolean zebra = false;
            for (DepartamentoDTO d: lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(d.getNomsucursal(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(d.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getNivel()), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(d.getDepartamento_padre(), ReporteUtil.fuenteTablaData(), fondo));
                zebra= !zebra;
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
