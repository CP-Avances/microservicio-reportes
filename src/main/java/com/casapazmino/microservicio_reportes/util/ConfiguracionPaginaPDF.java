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

    // 🔵 Constructor original con 3 parámetros
    public ConfiguracionPaginaPDF(String nombreUsuario, String textoMarcaAgua, String colorHexPrincipal) {
        this.nombreUsuario = nombreUsuario;
        this.textoMarcaAgua = textoMarcaAgua;
        this.colorPrincipal = ReporteUtil.convertirHexAColor(colorHexPrincipal);
    }

    // 🟢 Nuevo constructor adicional con 2 parámetros
    public ConfiguracionPaginaPDF(String nombreUsuario, String textoMarcaAgua) {
        this(nombreUsuario, textoMarcaAgua, "#C8C8FF"); // Color predeterminado: azul claro
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
        Font fuentePie = FontFactory.getFont(FontFactory.HELVETICA, 9, Font.NORMAL, new Color(120, 120, 120));

        String fecha = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        String hora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
        String textoIzquierdo = "Fecha: " + fecha + "  Hora: " + hora;
        String textoDerecho = "© Pag " + writer.getPageNumber();

        ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(textoIzquierdo, fuentePie),
                document.left(), document.bottom() - 10, 0);
        ColumnText.showTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(textoDerecho, fuentePie),
                document.right(), document.bottom() - 10, 0);

        Font fuenteEncabezado = FontFactory.getFont(FontFactory.HELVETICA, 9, Font.NORMAL, new Color(120, 120, 120));
        Phrase encabezado = new Phrase("Impreso por: " + nombreUsuario, fuenteEncabezado);
        ColumnText.showTextAligned(canvas, Element.ALIGN_RIGHT, encabezado, document.right(), document.top() + 10, 0);
    }
}
