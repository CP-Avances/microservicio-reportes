package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Rol.FuncionDTO;
import com.casapazmino.microservicio_reportes.model.Rol.ReporteRolesRequest;
import com.casapazmino.microservicio_reportes.model.Rol.RolDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteRolesService {

    //Recibe un objeto ReporteRolesRquest
    public byte[] generarReporteRolesPDF(ReporteRolesRequest request) {
        try {

            //Aqui se guardara el pdf
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
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
                document.add(logo);
            }

            // Empresa y Título
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("PERMISOS O FUNCIONALIDADES DEL ROL"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            for (RolDTO rol : request.getRoles()) {
                // Título del Rol
                PdfPTable encabezado = new PdfPTable(1);
                encabezado.setWidthPercentage(100);
                encabezado.setSpacingBefore(10f);

                PdfPCell celdaRol = new PdfPCell(new Phrase("ROL: " + rol.getNombre(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9)));
                celdaRol.setBackgroundColor(colorPrincipal);
                celdaRol.setPadding(5f);
                encabezado.addCell(celdaRol);
                document.add(encabezado);

                // Subtítulo
                PdfPTable subtitulo = new PdfPTable(1);
                subtitulo.setWidthPercentage(100);

                PdfPCell celdaTitulo = new PdfPCell(new Phrase("FUNCIONES DEL SISTEMA ASIGNADAS", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9)));
                celdaTitulo.setBackgroundColor(colorSecundario);
                celdaTitulo.setHorizontalAlignment(Element.ALIGN_CENTER);
                celdaTitulo.setPadding(5f);
                subtitulo.addCell(celdaTitulo);
                document.add(subtitulo);

                // Tabla
                PdfPTable tabla = new PdfPTable(5);
                tabla.setWidthPercentage(100);
                tabla.setWidths(new float[]{3, 4, 4, 2, 2});
                tabla.setSpacingBefore(5f);

                // Encabezados
                String[] headers = {"PÁGINA", "FUNCIÓN", "MÓDULO", "APLICACIÓN WEB", "APLICACIÓN MÓVIL"};
                for (String h : headers) {
                    tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), colorSecundario));
                }

                // Filas con efecto zebra, se agrega las funciones de los roles
                boolean zebra = false;
                for (FuncionDTO f : rol.getFunciones()) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    zebra = !zebra;

                    tabla.addCell(ReporteUtil.crearCelda(f.getPagina(), ReporteUtil.fuenteTexto(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(f.getAccion(), ReporteUtil.fuenteTexto(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(transformarModulo(f.getNombre_modulo()), ReporteUtil.fuenteTexto(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(f.isMovil() ? "" : "Sí", ReporteUtil.fuenteTexto(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(f.isMovil() ? "Sí" : "", ReporteUtil.fuenteTexto(), fondo));
                }
                document.add(tabla);
            }

            document.close();
            writer.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    //Metodo para poner los nombre de los modulos
    private String transformarModulo(String nombreModulo) {
        if (nombreModulo == null) return "";
        switch (nombreModulo) {
            case "permisos": return "Módulo de Permisos";
            case "vacaciones": return "Módulo de Vacaciones";
            case "horas_extras": return "Módulo de Horas Extras";
            case "alimentacion": return "Módulo de Alimentación";
            case "acciones_personal": return "Módulo de Acciones de Personal";
            case "geolocalizacion": return "Módulo de Geolocalización";
            case "timbre_virtual": return "Módulo de Timbre Virtual";
            case "reloj_virtual": return "Aplicación Móvil";
            case "aprobar": return "Aprobaciones Solicitudes";
            default: return nombreModulo;
        }
    }
}
