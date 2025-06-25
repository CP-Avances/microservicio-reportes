package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteRolesRequest;
import com.casapazmino.microservicio_reportes.model.RolDTO;
import com.casapazmino.microservicio_reportes.model.FuncionDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.List;

@Service
public class ReporteRolesService {

    public byte[] generarReporteRolesPDF(ReporteRolesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(request.getUsuario(), request.getFraseMarcaAgua()));
            document.open();

            // LOGO
            if (request.getLogoBase64() != null && request.getLogoBase64().contains("base64,")) {
                String base64Image = request.getLogoBase64().split(",")[1];
                byte[] imageBytes = Base64.getDecoder().decode(base64Image);
                Image logo = Image.getInstance(imageBytes);
                logo.scaleAbsolute(100, 50);
                logo.setAlignment(Image.LEFT);
                document.add(logo);
            }

            // EMPRESA
            Paragraph empresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingAfter(5f);
            empresa.setSpacingBefore(-30f);
            document.add(empresa);

            // TÍTULO
            Paragraph titulo = new Paragraph("PERMISOS O FUNCIONALIDADES DEL ROL", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // COLORES desde frontend
            Color colorPrincipal = Color.decode(request.getColorPrincipal());   // para encabezado de ROL
            Color colorSecundario = Color.decode(request.getColorSecundario()); // para encabezados de tabla y subtítulos
            Color colorZebra = new Color(229, 231, 233); // gris claro

            // ROLES
            for (RolDTO rol : request.getRoles()) {

                // Encabezado del ROL
                PdfPTable encabezado = new PdfPTable(1);
                encabezado.setWidthPercentage(100);
                encabezado.setSpacingBefore(10f);

                PdfPCell celdaRol = new PdfPCell(new Phrase("ROL: " + rol.getNombre(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9)));
                celdaRol.setBackgroundColor(colorPrincipal);
                celdaRol.setPadding(5f);
                encabezado.addCell(celdaRol);
                document.add(encabezado);

                // Subtítulo
                PdfPTable subtitulo = new PdfPTable(1);
                subtitulo.setWidthPercentage(100);

                PdfPCell celdaTitulo = new PdfPCell(new Phrase("FUNCIONES DEL SISTEMA ASIGNADAS", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9)));
                celdaTitulo.setBackgroundColor(colorSecundario);
                celdaTitulo.setHorizontalAlignment(Element.ALIGN_CENTER);
                celdaTitulo.setPadding(5f);
                subtitulo.addCell(celdaTitulo);
                document.add(subtitulo);

                // Tabla
                PdfPTable tabla = new PdfPTable(5);
                tabla.setWidthPercentage(100);
                tabla.setWidths(new float[]{3, 4, 4, 2, 2});
                tabla.setSpacingBefore(5f);

                // Encabezados
                String[] headers = {"PÁGINA", "FUNCIÓN", "MÓDULO", "APLICACIÓN WEB", "APLICACIÓN MÓVIL"};
                for (String h : headers) {
                    PdfPCell header = new PdfPCell(new Phrase(h, FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8)));
                    header.setBackgroundColor(colorSecundario);
                    header.setHorizontalAlignment(Element.ALIGN_CENTER);
                    header.setPadding(5f);
                    tabla.addCell(header);
                }

                // Filas
                boolean zebra = false;
                for (FuncionDTO f : rol.getFunciones()) {
                    Color bg = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(createCell(f.getPagina(), bg));
                    tabla.addCell(createCell(f.getAccion(), bg));
                    tabla.addCell(createCell(transformarModulo(f.getNombre_modulo()), bg));
                    tabla.addCell(createCell(f.isMovil() ? "" : "Sí", bg));
                    tabla.addCell(createCell(f.isMovil() ? "Sí" : "", bg));
                    zebra = !zebra;
                }

                document.add(tabla);
            }

            document.close();
            writer.close();
            return baos.toByteArray();

        } catch (IOException | DocumentException e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell createCell(String texto, Color bg) {
        PdfPCell cell = new PdfPCell(new Phrase(texto != null ? texto : "", FontFactory.getFont(FontFactory.HELVETICA, 8)));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setBackgroundColor(bg);
        cell.setPadding(4f);
        return cell;
    }

    private String transformarModulo(String nombreModulo) {
        if (nombreModulo == null) return "";
        switch (nombreModulo) {
            case "permisos": return "Módulo de Permisos";
            case "vacaciones": return "Módulo de Vacaciones";
            case "horas_extras": return "Módulo de Horas Extras";
            case "alimentacion": return "Módulo de Alimentación";
            case "acciones_personal": return "Módulo de Acciones de Personal";
            case "geolocalizacion": return "Módulo de Geolocalización";
            case "timbre_virtual": return "Módulo de Timbre Virtual";
            case "reloj_virtual": return "Aplicación Móvil";
            case "aprobar": return "Aprobaciones Solicitudes";
            default: return nombreModulo;
        }
    }
}
