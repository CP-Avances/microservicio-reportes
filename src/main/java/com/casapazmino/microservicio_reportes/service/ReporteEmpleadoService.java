package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.EmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.ReporteEmpleadosRequest;
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

    public byte[] generarReporteEmpleadosPDF(ReporteEmpleadosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate());
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
                logo.scaleToFit(120, 120);
                document.add(logo);
            }

            // Empresa
            Paragraph empresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-40f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            Paragraph titulo = new Paragraph("Lista de Empleados", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Tabla
            PdfPTable tabla = new PdfPTable(11);
            tabla.setWidthPercentage(100);
            tabla.setWidths(new float[]{2, 4, 4, 3, 5, 3, 3, 4, 3, 3, 4});
            tabla.setSpacingBefore(10f);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Encabezados
            String[] headers = {
                "Código", "Nombre", "Identificación", "Fecha Nacimiento", "Correo",
                "Género", "Estado Civil", "Domicilio", "Teléfono", "Estado", "Nacionalidad"
            };

            for (String encabezado : headers) {
                tabla.addCell(crearCelda(encabezado, colorPrincipal));
            }

            List<EmpleadoDTO> empleados = request.getEmpleados();
            for (int i = 0; i < empleados.size(); i++) {
                EmpleadoDTO e = empleados.get(i);
                Color bgColor = (i % 2 == 0) ? colorZebra : null;

                tabla.addCell(ReporteUtil.crearCelda(e.getCodigo(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getNombreCompleto(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getIdentificacion(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getFechaNacimiento(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getCorreo(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getGenero(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getEstadoCivil(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getDomicilio(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getTelefono(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getEstadoTexto(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(e.getNacionalidad(), ReporteUtil.fuenteTexto(), bgColor));
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell crearCelda(String texto, Color bgColor) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, ReporteUtil.fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setBackgroundColor(bgColor);
        return celda;
    }
}
