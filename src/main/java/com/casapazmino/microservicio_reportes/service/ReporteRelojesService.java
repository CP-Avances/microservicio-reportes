package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Dispositivo.RelojDTO;
import com.casapazmino.microservicio_reportes.model.Dispositivo.ReporteRelojesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteRelojesService {

    public byte[] generarReportePDF(ReporteRelojesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 50, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                logo.scaleAbsolute(100, 50);
                logo.setAlignment(Image.LEFT);
                document.add(logo);
            }

            // Empresa y título (usando métodos utilitarios)
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DISPOSITIVOS"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(15);
            tabla.setWidthPercentage(100);
            tabla.setWidths(new float[]{
                    2, 3, 3, 3.5f, 3, 3.5f, 2.5f, 2, 2.5f, 2.5f, 3, 3, 3, 3, 4
            });
            tabla.setSpacingBefore(5f);

            String[] headers = {
                    "Código", "Empresa", "Ciudad", "Establecimiento", "Departamento", "Nombre", "IP",
                    "Puerto", "Marca", "Modelo", "Serie", "Mac", "ID Fabricación", "Fabricante", "Zona Horaria"
            };

            for (String h : headers) {
                tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), colorPrincipal));
            }

            boolean zebra = false;
            List<RelojDTO> lista = request.getRelojes();

            for (RelojDTO r : lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                zebra = !zebra;

                tabla.addCell(ReporteUtil.crearCelda(r.getCodigo(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomempresa(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomciudad(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomsucursal(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomdepar(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNombre(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getIp(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getPuerto()), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getMarca(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getModelo(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getSerie(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getMac(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getIdFabricacion(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getFabricante(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(
                        r.getZonaHorariaDispositivo() + " (" + r.getFormatoGmtDispositivo() + ")",
                        ReporteUtil.fuenteTexto(), fondo));
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
