package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteAuditoria.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAuditoriaService {

    public byte[] generarReportePDF(ReporteAuditoriaRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("AUDITORÍA"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebra = ReporteUtil.colorZebraClaro();

            List<AuditoriaDTO> lista = request.getAuditorias();
            if (lista == null || lista.isEmpty()) {
                document.add(new Paragraph("No hay registros para mostrar.", ReporteUtil.fuenteTexto()));
                document.close();
                return baos.toByteArray();
            }

            // Cabecera: PLATAFORMA + N° Registros
            PdfPTable cabecera = new PdfPTable(2);
            cabecera.setWidthPercentage(100);
            cabecera.setWidths(new float[]{8, 2});
            cabecera.setSpacingBefore(10f);

            PdfPCell celdaPlataforma = new PdfPCell(new Phrase("PLATAFORMA: " + lista.get(0).getPlataforma(), ReporteUtil.fuenteTexto()));
            celdaPlataforma.setBackgroundColor(colorSecundario);
            celdaPlataforma.setBorder(Rectangle.TOP | Rectangle.LEFT | Rectangle.BOTTOM);
            celdaPlataforma.setPadding(5);
            cabecera.addCell(celdaPlataforma);

            PdfPCell celdaCantidad = new PdfPCell(new Phrase("N° Registros: " + lista.size(), ReporteUtil.fuenteTexto()));
            celdaCantidad.setBackgroundColor(colorSecundario);
            celdaCantidad.setBorder(Rectangle.TOP | Rectangle.RIGHT | Rectangle.BOTTOM);
            celdaCantidad.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaCantidad.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaCantidad.setPadding(5);
            cabecera.addCell(celdaCantidad);

            document.add(cabecera);

            // Tabla de auditoría
            PdfPTable tabla = new PdfPTable(10);
            tabla.setWidthPercentage(100);
            tabla.setSpacingBefore(5f);
            tabla.setWidths(new float[]{1f, 3f, 2f, 2.5f, 3f, 2f, 2f, 2f, 6f, 6f});

            // Encabezados
            String[] encabezados = {
                    "ITEM", "PLATAFORMA", "USUARIO", "IP", "NOMBRE TABLA", "ACCIÓN", "FECHA", "HORA", "DATOS ORIGINALES", "DATOS NUEVOS"
            };

            for (String enc : encabezados) {
                tabla.addCell(ReporteUtil.crearCelda(enc, ReporteUtil.fuenteEncabezado(), colorPrincipal));
            }

            // Cuerpo
            AtomicInteger index = new AtomicInteger(1);
            for (AuditoriaDTO a : lista) {
                Color fondo = (index.get() % 2 == 0) ? zebra : Color.WHITE;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(index.getAndIncrement()), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getPlataforma(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getUser_name(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getIp_address(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getTable_name(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getAction(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getFecha_hora_format(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getSolo_hora(), ReporteUtil.fuenteTexto(), fondo));

                PdfPCell celdaOriginal = ReporteUtil.crearCelda(a.getOriginal_data(), ReporteUtil.fuenteTexto(), fondo);
                celdaOriginal.setNoWrap(false);
                celdaOriginal.setMinimumHeight(20f);
                tabla.addCell(celdaOriginal);

                PdfPCell celdaNuevo = ReporteUtil.crearCelda(a.getNew_data(), ReporteUtil.fuenteTexto(), fondo);
                celdaNuevo.setNoWrap(false);
                celdaNuevo.setMinimumHeight(20f);
                tabla.addCell(celdaNuevo);
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
