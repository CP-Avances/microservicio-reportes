package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Sucursal.ReporteSucursalesRequest;
import com.casapazmino.microservicio_reportes.model.Sucursal.SucursalDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteSucursalesService {

    public byte[] generarReportePDF(ReporteSucursalesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
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

            // Empresa
            Paragraph empresa = new Paragraph(request.getEmpresa(),
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-30f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            Paragraph titulo = new Paragraph("LISTA DE SUCURSALES",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Tabla
            PdfPTable tabla = new PdfPTable(3);
            tabla.setWidthPercentage(80); // más amplio por tener 3 columnas
            tabla.setWidths(new float[]{1.5f, 4, 3});
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("SUCURSAL / ESTABLECIMIENTO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("CIUDAD", ReporteUtil.fuenteEncabezado(), colorPrincipal));

            // Filas con zebra
            boolean zebra = false;
            List<SucursalDTO> lista = request.getSucursales();
            for (SucursalDTO sucursal : lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(sucursal.getId()), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(sucursal.getNombre(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(sucursal.getDescripcion(), ReporteUtil.fuenteTexto(), fondo));
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
