package com.casapazmino.microservicio_reportes.util;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import java.awt.Color;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ConfiguracionPaginaPDF extends PdfPageEventHelper {

    private final String nombreUsuario;
    private final String textoMarcaAgua;
    private final Color colorPrincipal;

    private PdfTemplate total;

    public ConfiguracionPaginaPDF(String nombreUsuario, String textoMarcaAgua, String colorHexPrincipal) {
        this.nombreUsuario = nombreUsuario;
        this.textoMarcaAgua = textoMarcaAgua;
        this.colorPrincipal = ReporteUtil.convertirHexAColor(colorHexPrincipal);
    }

    public ConfiguracionPaginaPDF(String nombreUsuario, String textoMarcaAgua) {
        this(nombreUsuario, textoMarcaAgua, "#C8C8FF"); 
    }

    @Override
    public void onOpenDocument(PdfWriter writer, Document document) {
        total = writer.getDirectContent().createTemplate(30, 16);
    }

    @Override
    public void onStartPage(PdfWriter writer, Document document) {
        if (textoMarcaAgua != null && !textoMarcaAgua.isEmpty()) {
            PdfContentByte canvas = writer.getDirectContentUnder();
            Font fuenteMarca = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 160, Font.NORMAL, colorPrincipal);
            Phrase marcaAgua = new Phrase(textoMarcaAgua, fuenteMarca);

            float x = document.getPageSize().getWidth() / 2 + 40;
            float y = (document.top() + document.bottom()) / 2;

            ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, marcaAgua, x, y, 50);
        }
    }

    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        PdfContentByte canvas = writer.getDirectContent();
        BaseFont baseFont;
        try {
            baseFont = BaseFont.createFont();
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        Font fuentePie = new Font(baseFont, 9, Font.NORMAL, new Color(120, 120, 120));

        // Fecha y hora
        String fecha = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        String hora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
        String textoIzquierdo = "Fecha: " + fecha + "  Hora: " + hora;
        ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(textoIzquierdo, fuentePie),
                document.left(), document.bottom() - 10, 0);

        // Página actual
        int pageNum = writer.getPageNumber();
        String texto = "© Pag " + pageNum + " de ";

        float textBase = document.bottom() - 10;
        float textSize = baseFont.getWidthPoint(texto, 9);
        float adjustX = document.right() - textSize - 20;

        // Escribir "© Pag X of "
        canvas.beginText();
        canvas.setFontAndSize(baseFont, 9);
        canvas.setTextMatrix(adjustX, textBase);
        canvas.setColorFill(new Color(120, 120, 120));
        canvas.showText(texto);
        canvas.endText();

        // Insertar plantilla del total justo después del texto
        canvas.addTemplate(total, adjustX + textSize, textBase);

        // Encabezado: usuario
        Font fuenteEncabezado = new Font(baseFont, 9, Font.NORMAL, new Color(120, 120, 120));
        Phrase encabezado = new Phrase("Impreso por: " + nombreUsuario, fuenteEncabezado);
        ColumnText.showTextAligned(canvas, Element.ALIGN_RIGHT, encabezado, document.right(), document.top() + 10, 0);
    }

    @Override
    public void onCloseDocument(PdfWriter writer, Document document) {
        try {
            BaseFont baseFont = BaseFont.createFont();
            total.beginText();
            total.setFontAndSize(baseFont, 9);
            total.setTextMatrix(0, 0);
            total.setColorFill(new Color(120, 120, 120));
            total.showText(String.valueOf(writer.getPageNumber() - 1));
            total.endText();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
