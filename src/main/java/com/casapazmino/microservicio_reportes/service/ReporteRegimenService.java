package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Regimen.RegimenDTO;
import com.casapazmino.microservicio_reportes.model.Regimen.ReporteRegimenRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;


import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteRegimenService {

    public byte[] generarReporteRegimenPDF(ReporteRegimenRequest request) throws Exception {
        Document document = new Document(PageSize.A4.rotate(), 36, 36, 90, 36);
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        PdfWriter writer = PdfWriter.getInstance(document, out);
        writer.setPageEvent(new ConfiguracionPaginaPDF(
                request.getUsuario(),
                request.getFraseMarcaAgua(),
                request.getColorPrincipal()
        ));

        document.open();

        // Agregar logo, empresa y título (título fijo para este módulo)
        if (request.getLogoBase64() != null) {
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);
        }

        Paragraph nombreEmpresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
        nombreEmpresa.setAlignment(Element.ALIGN_CENTER);
        nombreEmpresa.setSpacingAfter(5f);
        nombreEmpresa.setSpacingBefore(-30f);
        document.add(nombreEmpresa);

        Paragraph titulo = new Paragraph("RÉGIMEN LABORAL", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
        titulo.setAlignment(Element.ALIGN_CENTER);
        titulo.setSpacingAfter(10f);
        document.add(titulo);

        // Colores
        Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

        // Tabla
        PdfPTable tabla = new PdfPTable(11);
        tabla.setWidthPercentage(100);
        tabla.setSpacingBefore(10f);
        tabla.setWidths(new float[]{2, 6, 4, 3, 3, 3, 3, 3, 4, 4, 3});

        Font fontHeader = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9, Color.WHITE);
        Font fontBody = ReporteUtil.fuenteTexto();

        String[] encabezados = {
                "CÓDIGO", "DESCRIPCIÓN", "PAÍS", "MESES PERIODO", "DÍAS POR MES", "VACACIONES POR AÑO",
                "DÍAS LIBRES", "DÍAS CALENDARIO", "DÍAS MÁXIMOS ACUMULABLES", "AÑOS PARA ANTIGUEDAD", "DÍAS DE INCREMENTO"
        };

        for (String tituloCol : encabezados) {
            PdfPCell celda = new PdfPCell(new Phrase(tituloCol, fontHeader));
            celda.setBackgroundColor(colorPrincipal);
            celda.setHorizontalAlignment(Element.ALIGN_CENTER);
            celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celda.setPadding(4f);
            tabla.addCell(celda);
        }

        List<RegimenDTO> lista = request.getRegimen();
        for (int i = 0; i < lista.size(); i++) {
            RegimenDTO r = lista.get(i);
            Color fondo = (i % 2 == 0) ? colorSecundario : Color.WHITE;

            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getId()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(r.getDescripcion(), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(r.getPais(), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getMes_periodo()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getDias_mes()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getVacacion_dias_laboral()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getVacacion_dias_libre()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getVacacion_dias_calendario()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getDias_maximo_acumulacion()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getAnio_antiguedad()), fontBody, fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getDias_antiguedad()), fontBody, fondo));
        }

        document.add(tabla);
        document.close();

        return out.toByteArray();
    }
}
