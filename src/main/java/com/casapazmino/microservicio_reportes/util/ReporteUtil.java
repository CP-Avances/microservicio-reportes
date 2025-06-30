package com.casapazmino.microservicio_reportes.util;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import java.awt.Color;
import java.util.Base64;

public class ReporteUtil {

    // Metodo para convertir código HEX (como "#E5E7E9") a objeto Color
    public static Color convertirHexAColor(String hex) {
        try {
            return Color.decode(hex);
        } catch (Exception e) {
            return Color.LIGHT_GRAY; 
        }
    }

    // Metodo para convertir logo base64 a Image
    public static Image obtenerLogo(String base64String) throws Exception {
        if (base64String != null && base64String.contains("base64,")) {
            String base64 = base64String.split(",")[1];
            byte[] imageBytes = Base64.getDecoder().decode(base64);
            Image logo = Image.getInstance(imageBytes);
            logo.scaleAbsolute(100, 100);
            logo.setAlignment(Image.LEFT);
            return logo;
        }
        return null; //Se podria ingresar un logo por defecto
    }

    // Metodo para crear celda personalizada con fondo y alineación
    public static PdfPCell crearCelda(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    // Metodo para fuente de texto
    public static Font fuenteTexto() {
        return FontFactory.getFont(FontFactory.HELVETICA, 8);
    }

    // Metodo para fuente para encabezados
    public static Font fuenteEncabezado() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9);
    }

    // Celda alineada al centro
    public static PdfPCell celdaCentro(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    // Celda alineada a la izquierda
    public static PdfPCell celdaIzquierda(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    // Celda para info de empleado gris claro
    public static PdfPCell celdaInfoEmpleado(String texto) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuenteTexto()));
        celda.setBackgroundColor(new Color(227, 227, 227)); // gris claro
        celda.setPadding(5);
        return celda;
    }

    // Celda de encabezado con color de fondo
    public static PdfPCell celdaEncabezado(String texto, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setBackgroundColor(fondo);
        celda.setPadding(5);
        return celda;
    }

    // Celda de encabezado para número de día
    public static PdfPCell celdaEncabezadoDia(int dia) {
        PdfPCell celda = new PdfPCell(new Phrase(String.format("%02d", dia), fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4);
        return celda;
    }

    public static PdfPCell celdaCentro(String texto) {
        return celdaCentro(texto, fuenteTexto());
    }

    public static PdfPCell celdaIzquierda(String texto) {
        return celdaIzquierda(texto, fuenteTexto());
    }

    public static PdfPCell celdaNomenclatura(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        celda.setMinimumHeight(20f); // altura constante
        celda.setNoWrap(true); // evita salto de línea
        return celda;
    }

    public static PdfPCell celdaNomenclaturaDescripcion(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        celda.setMinimumHeight(20f); // igual altura que nombre
        celda.setNoWrap(true); // evita salto de línea
        return celda;
    }

}