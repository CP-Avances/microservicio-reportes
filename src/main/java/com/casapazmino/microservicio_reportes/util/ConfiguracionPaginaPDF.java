package com.casapazmino.microservicio_reportes.util;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

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
    }

    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        PdfContentByte canvas = writer.getDirectContent();

        if (textoMarcaAgua != null && !textoMarcaAgua.isEmpty()) {
            PdfGState gstate = new PdfGState();
            boolean landscape = document.getPageSize().getWidth() > document.getPageSize().getHeight();
            float angulo = landscape ? 30f : 55f; 
            float fontSize = landscape ? 180f : 160f; 

            gstate.setFillOpacity(0.1f);
            canvas.saveState();
            canvas.setGState(gstate);
            Font fuenteMarca = FontFactory.getFont(FontFactory.HELVETICA_BOLD, fontSize, Font.NORMAL, Color.blue);
            Phrase marcaAgua = new Phrase(textoMarcaAgua, fuenteMarca);

            float x = document.getPageSize().getWidth() / 2 + 40;
            float y = (document.top() + document.bottom()) / 2;

            ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, marcaAgua, x, y, angulo);
            canvas.restoreState();
        }

        BaseFont baseFont;
        try {
            baseFont = BaseFont.createFont();
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        float pageWidth = document.getPageSize().getWidth();
        float pageHeight = document.getPageSize().getHeight();

        Font fuentePie = new Font(baseFont, 9, Font.NORMAL, new Color(120, 120, 120));

        PdfGState gstate = new PdfGState();
        gstate.setFillOpacity(0.5f); 
        canvas.saveState();
        canvas.setGState(gstate);

        String fecha = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        String hora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
        String textoIzquierdo = "Fecha: " + fecha + "  Hora: " + hora;

        ColumnText.showTextAligned(
                canvas,
                Element.ALIGN_LEFT,
                new Phrase(textoIzquierdo, fuentePie),
                20, 
                20,
                0);

        int pageNum = writer.getPageNumber();
        String textoPagina = "© Pag " + pageNum + " de ";
        float textSize = baseFont.getWidthPoint(textoPagina, 9);

        float xPagina = pageWidth - textSize - 30;
        float yPagina = 20;

        canvas.beginText();
        canvas.setFontAndSize(baseFont, 9);
        canvas.setTextMatrix(xPagina, yPagina);
        canvas.setColorFill(new Color(120, 120, 120));
        canvas.showText(textoPagina);
        canvas.endText();

        canvas.addTemplate(total, xPagina + textSize, yPagina);

        canvas.restoreState();

        Font fuenteEncabezado = new Font(baseFont, 9, Font.NORMAL, new Color(120, 120, 120));
        Phrase encabezado = new Phrase("Impreso por: " + nombreUsuario, fuenteEncabezado);

        gstate.setFillOpacity(0.5f); 

        canvas.saveState();
        canvas.setGState(gstate);

        ColumnText.showTextAligned(
                canvas,
                Element.ALIGN_RIGHT,
                encabezado,
                pageWidth - 15, 
                pageHeight - 20, 
                0);

        canvas.restoreState(); 

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
