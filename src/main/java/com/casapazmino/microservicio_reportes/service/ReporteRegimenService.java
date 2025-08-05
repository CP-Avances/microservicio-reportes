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

        // Logo
        Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
        if (logo != null) {
            document.add(logo);
        }

        // Empresa y Título
        document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
        document.add(ReporteUtil.crearTituloReporte("RÉGIMEN LABORAL"));

        // Colores
        Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        Color colorZebra = ReporteUtil.colorZebraClaro();

        // Tabla
        PdfPTable tabla = new PdfPTable(11);
        tabla.setWidthPercentage(100);
        tabla.setSpacingBefore(10f);
        tabla.setWidths(new float[]{2, 6, 4, 3, 3, 3, 3, 3, 4, 4, 3});

        String[] encabezados = {
                "CÓDIGO", "DESCRIPCIÓN", "PAÍS", "MESES PERIODO", "DÍAS POR MES", "VACACIONES POR AÑO",
                "DÍAS LIBRES", "DÍAS CALENDARIO", "DÍAS MÁXIMOS ACUMULABLES", "AÑOS PARA ANTIGUEDAD", "DÍAS DE INCREMENTO"
        };

        for (String tituloCol : encabezados) {
            tabla.addCell(ReporteUtil.crearCelda(tituloCol, ReporteUtil.fuenteEncabezado(), colorPrincipal));
        }

        List<RegimenDTO> lista = request.getRegimen();
        for (int i = 0; i < lista.size(); i++) {
            RegimenDTO r = lista.get(i);
            Color fondo = (i % 2 == 0) ? colorZebra : null;

            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getId()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(r.getDescripcion(), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(r.getPais(), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getMes_periodo()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getDias_mes()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getVacacion_dias_laboral()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getVacacion_dias_libre()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getVacacion_dias_calendario()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getDias_maximo_acumulacion()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getAnio_antiguedad()), ReporteUtil.fuenteTexto(), fondo));
            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getDias_antiguedad()), ReporteUtil.fuenteTexto(), fondo));
        }

        document.add(tabla);
        document.close();

        return out.toByteArray();
    }
}
