package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Empleado.EmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.Empleado.ReporteEmpleadosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteEmpleadoService {

    //METODO QU GENERA EL REPORTE PDF
    public byte[] generarReporteEmpleadosPDF(ReporteEmpleadosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4.rotate());
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
            document.add(ReporteUtil.crearTituloReporte("Lista de Empleados"));

            //TABLA
            PdfPTable tabla = new PdfPTable(11);
            tabla.setWidthPercentage(100);
            tabla.setWidths(new float[]{1.8f, 5.5f, 3.7f, 3, 7.3f, 2.7f, 3, 3, 3, 2, 3});
            tabla.setSpacingBefore(10f);

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            //ENCABEZADOS
            String[] headers = {
                "Código", "Nombre", "Identificación", "Fecha Nacimiento", "Correo",
                "Género", "Estado Civil", "Domicilio", "Teléfono", "Estado", "Nacionalidad"
            };

            for (String encabezado : headers) {
                tabla.addCell(ReporteUtil.crearCelda(encabezado, ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            }

            List<EmpleadoDTO> empleados = request.getEmpleados();
            boolean zebra = false;
            for (EmpleadoDTO e: empleados) {
                Color bgColor = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(e.getCodigo(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getNombreCompleto(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getIdentificacion(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getFechaNacimiento(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getCorreo(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getGenero(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getEstadoCivil(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getDomicilio(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getTelefono(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getEstadoTexto(), ReporteUtil.fuenteTablaData(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getNacionalidad(), ReporteUtil.fuenteTablaData(), bgColor));
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
