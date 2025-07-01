package com.casapazmino.microservicio_reportes.util;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import java.awt.Color;
import java.util.Base64;

public class ReporteUtil {

    // Convertir código HEX (como "#E5E7E9") a objeto Color
    public static Color convertirHexAColor(String hex) {
        try {
            return Color.decode(hex);
        } catch (Exception e) {
            return Color.LIGHT_GRAY;
        }
    }

    // Convertir logo base64 a Image
    public static Image obtenerLogo(String base64String) throws Exception {
        if (base64String != null && base64String.contains("base64,")) {
            String base64 = base64String.split(",")[1];
            byte[] imageBytes = Base64.getDecoder().decode(base64);
            Image logo = Image.getInstance(imageBytes);
            aplicarEstiloLogo(logo);
            return logo;
        }
        return null;
    }

    // Aplicar tamaño y alineación al logo
    public static void aplicarEstiloLogo(Image logo) {
        logo.scaleToFit(120, 120);
        logo.setAlignment(Image.LEFT);
    }

    // Crear celda con texto, fuente y fondo personalizado
    public static PdfPCell crearCelda(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    // Fuente de texto general
    public static Font fuenteTexto() {
        return FontFactory.getFont(FontFactory.HELVETICA, 8);
    }

    // Fuente para encabezados
    public static Font fuenteEncabezado() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9);
    }

    // Fuente para título de empresa
    public static Font fuenteTituloEmpresa() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14);
    }

    // Fuente para título del reporte
    public static Font fuenteTituloReporte() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
    }

    // Crear título de empresa
    public static Paragraph crearTituloEmpresa(String texto) {
        Paragraph p = new Paragraph(texto, fuenteTituloEmpresa());
        p.setAlignment(Element.ALIGN_CENTER);
        p.setSpacingBefore(-30f);
        p.setSpacingAfter(5f);
        return p;
    }

    // Crear título del reporte
    public static Paragraph crearTituloReporte(String texto) {
        Paragraph p = new Paragraph(texto, fuenteTituloReporte());
        p.setAlignment(Element.ALIGN_CENTER);
        p.setSpacingAfter(10f);
        return p;
    }

    // Color zebra claro reutilizable
    public static Color colorZebraClaro() {
        return new Color(204, 209, 209); // #CCD1D1
    }

    // Celda alineada al centro
    public static PdfPCell celdaCentro(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    public static PdfPCell celdaCentro(String texto) {
        return celdaCentro(texto, fuenteTexto());
    }

    // Celda alineada a la izquierda
    public static PdfPCell celdaIzquierda(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    public static PdfPCell celdaIzquierda(String texto) {
        return celdaIzquierda(texto, fuenteTexto());
    }

    // Celda gris para info de empleado
    public static PdfPCell celdaInfoEmpleado(String texto) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuenteTexto()));
        celda.setBackgroundColor(new Color(227, 227, 227));
        celda.setPadding(5);
        return celda;
    }

    // Celda de encabezado con fondo
    public static PdfPCell celdaEncabezado(String texto, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setBackgroundColor(fondo);
        celda.setPadding(5);
        return celda;
    }

    // Celda especial para número de día (horario)
    public static PdfPCell celdaEncabezadoDia(int dia) {
        PdfPCell celda = new PdfPCell(new Phrase(String.format("%02d", dia), fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4);
        return celda;
    }

    // Celdas para nomenclatura
    public static PdfPCell celdaNomenclatura(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        celda.setMinimumHeight(20f);
        celda.setNoWrap(true);
        return celda;
    }

    public static PdfPCell celdaNomenclaturaDescripcion(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        celda.setMinimumHeight(20f);
        celda.setNoWrap(true);
        return celda;
    }
}
