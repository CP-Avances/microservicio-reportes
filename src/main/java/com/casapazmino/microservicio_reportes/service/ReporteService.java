package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteGenerosRequest;
import com.casapazmino.microservicio_reportes.util.ReporteGenerosPageEvent;
import com.casapazmino.microservicio_reportes.model.GeneroDTO;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.List;
import java.awt.Color;


@Service
public class ReporteService {

    public byte[] generarReporteGenerosPDF(ReporteGenerosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ReporteGenerosPageEvent(request.getUsuario(), request.getFraseMarcaAgua()));

            document.open();

            //LOGO
            if (request.getLogoBase64() != null && request.getLogoBase64().contains("base64,")) {
                String base64Image = request.getLogoBase64().split(",")[1];
                byte[] imageBytes = Base64.getDecoder().decode(base64Image);
                Image logo = Image.getInstance(imageBytes);
                logo.scaleAbsolute(100, 50);
                logo.setAlignment(Image.LEFT);
                document.add(logo);
            }

            //NOMBRE EMPRESA
            Paragraph nombreEmpresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            nombreEmpresa.setAlignment(Element.ALIGN_CENTER);
            document.add(nombreEmpresa);

            //TITULO DE REPORTE
            Paragraph titulo = new Paragraph("LISTA DE GÉNEROS", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            document.add(titulo);


            //TABLA
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(80);
            tabla.setSpacingBefore(10f);
            tabla.setSpacingAfter(10f);
            tabla.setHorizontalAlignment(Element.ALIGN_CENTER);

            //ENCABEZADO
            PdfPCell header1 = new PdfPCell(new Phrase("CÓDIGO"));
            header1.setHorizontalAlignment(Element.ALIGN_CENTER);
            header1.setBackgroundColor(new Color(204, 204, 204));
            tabla.addCell(header1);

            PdfPCell header2 = new PdfPCell(new Phrase("GÉNERO"));
            header2.setHorizontalAlignment(Element.ALIGN_CENTER);
            header2.setBackgroundColor(new Color(204, 204, 204));
            tabla.addCell(header2);

            List<GeneroDTO> generos = request.getGeneros();
            for (GeneroDTO genero : generos) {
  
                PdfPCell celda1 = new PdfPCell(new Phrase(String.valueOf(genero.getId())));
                celda1.setHorizontalAlignment(Element.ALIGN_CENTER);
                tabla.addCell(celda1);
            
                PdfPCell celda2 = new PdfPCell(new Phrase(genero.getGenero()));
                celda2.setHorizontalAlignment(Element.ALIGN_CENTER);
                tabla.addCell(celda2);

            }
            
            
            document.add(tabla);

            document.close();
            writer.close();

            return baos.toByteArray();

        } catch (IOException | DocumentException e) {
            e.printStackTrace();
            return null;
        }
    }
}
