package com.casapazmino.microservicio_reportes.util;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import java.awt.Color;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ReporteGenerosPageEvent extends PdfPageEventHelper {

    private final String usuario;
    private final String fraseMarcaAgua;

    public ReporteGenerosPageEvent(String usuario, String fraseMarcaAgua) {
        this.usuario = usuario;
        this.fraseMarcaAgua = fraseMarcaAgua;
    }

    //MARCA DE AGUA
    @Override
    public void onStartPage(PdfWriter writer, Document document) {
        if (fraseMarcaAgua != null && !fraseMarcaAgua.isEmpty()) {
            PdfContentByte cb = writer.getDirectContentUnder();
            Font watermarkFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 160, Font.NORMAL, new Color(230, 230, 250));
            Phrase watermark = new Phrase(fraseMarcaAgua, watermarkFont);
    
            float x = document.getPageSize().getWidth() / 2+40;
            float y = (document.top() + document.bottom()) / 2;
    
            ColumnText.showTextAligned(cb, Element.ALIGN_CENTER, watermark, x, y, 50); 
        }
    }
    

    //ENCABEZADO Y PIE DE PAGINA
    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        PdfContentByte cb = writer.getDirectContent();
        Font fontFooter = FontFactory.getFont(FontFactory.HELVETICA, 9, Font.NORMAL, new Color(120, 120, 120));

        //FECHA, HORA Y NUMERO DE PAGINA
        String fecha = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        String hora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
        String pieIzq = "Fecha: " + fecha + "  Hora: " + hora;
        String pieDer = "© Pag " + writer.getPageNumber();

        ColumnText.showTextAligned(cb, Element.ALIGN_LEFT, new Phrase(pieIzq, fontFooter),
                document.left(), document.bottom() - 10, 0);
        ColumnText.showTextAligned(cb, Element.ALIGN_RIGHT, new Phrase(pieDer, fontFooter),
                document.right(), document.bottom() - 10, 0);

        //IMPRESO POR : USUARIO QUE IMPRIME
        Font fontHeader = FontFactory.getFont(FontFactory.HELVETICA, 9, Font.NORMAL, new Color(120, 120, 120));
        Phrase header = new Phrase("Impreso por: " + usuario, fontHeader);
        ColumnText.showTextAligned(cb, Element.ALIGN_RIGHT, header, document.right(), document.top() + 10, 0);
    }
}
