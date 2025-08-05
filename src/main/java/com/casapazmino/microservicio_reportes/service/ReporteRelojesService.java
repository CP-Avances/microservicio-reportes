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

    // METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteRelojesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate());
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // LOGO DE EMPRESA
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // TITULO DE EMPRESA
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

            // TITULO DE REPORTE
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DISPOSITIVOS"));

            // COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // TABLA
            PdfPTable tabla = new PdfPTable(15);
            tabla.setWidthPercentage(100);
            tabla.setWidths(new float[] {
                    2.5f, 4, 2.2f, 4, 3.7f, 2.5f, 3, 2, 2, 2.5f, 3, 1.5f, 3.5f, 3, 4.5f
            });
            tabla.setSpacingBefore(10f);

            String[] headers = {
                    "Código", "Empresa", "Ciudad", "Establecimiento", "Departamento", "Nombre", "IP",
                    "Puerto", "Marca", "Modelo", "Serie", "Mac", "ID Fabricación", "Fabricante", "Zona Horaria"
            };

            // CABEZERAS DE LA TABLA
            for (String h : headers) {
                tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            }

            // FILAS DE LA TABLA (CUERPO)
            boolean zebra = false;
            List<RelojDTO> lista = request.getRelojes();
            for (RelojDTO r : lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                zebra = !zebra;

                tabla.addCell(ReporteUtil.crearCelda(r.getCodigo(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomempresa(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomciudad(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomsucursal(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNomdepar(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getNombre(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getIp(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(
                        ReporteUtil.crearCelda(String.valueOf(r.getPuerto()), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getMarca(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getModelo(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getSerie(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getMac(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getIdFabricacion(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getFabricante(), fuenteAuxiliarTabla(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(r.getZonaHorariaDispositivo() + " (" + r.getFormatoGmtDispositivo() + ")",fuenteAuxiliarTabla(), fondo));
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public Font fuenteAuxiliarTabla() {
        return FontFactory.getFont(FontFactory.HELVETICA, 7);
    }
}
