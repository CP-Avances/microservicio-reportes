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

@Service
public class ReporteRolesService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteRolesPDF(ReporteRolesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            //LOGO
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            //TITULO DE EMPRESA (EJM. CASA PAZMIÑO S.A.)
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            
            //TITULO DE REPORTE (EJM. REPORTE ATRASOS)
            document.add(ReporteUtil.crearTituloReporte("PERMISOS O FUNCIONALIDADES DEL ROL"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            for (RolDTO rol : request.getRoles()) {
                // TABLA ENCABEZADO ROL
                PdfPTable encabezado = new PdfPTable(1);
                encabezado.setWidthPercentage(100);
                encabezado.setSpacingBefore(10f);

                PdfPCell celdaRol = new PdfPCell(new Phrase("ROL: " + rol.getNombre(), ReporteUtil.fuenteEncabezado()));
                celdaRol.setBackgroundColor(colorPrincipal);
                celdaRol.setPadding(3f);
                encabezado.addCell(celdaRol);

                encabezado.setSpacingAfter(5f);
                document.add(encabezado);

                //TABLA DE SUBTITULO DE ESTE REPORTE (FUNCIONES DEL SISTEMA ASIGANDAS)
                PdfPTable subtitulo = new PdfPTable(1);
                subtitulo.setWidthPercentage(100);

                PdfPCell celdaTitulo = new PdfPCell(new Phrase("FUNCIONES DEL SISTEMA ASIGNADAS", ReporteUtil.fuenteEncabezadoTablaData()));
                celdaTitulo.setBackgroundColor(colorSecundario);
                celdaTitulo.setHorizontalAlignment(Element.ALIGN_CENTER);
                celdaTitulo.setPadding(3f);
                subtitulo.addCell(celdaTitulo);
                document.add(subtitulo);

                //TABLA DE INFORMACION DE LOS ROLES
                PdfPTable tabla = new PdfPTable(5);
                tabla.setWidthPercentage(100);
                tabla.setWidths(new float[]{3, 4, 4, 2, 2});
                tabla.setSpacingBefore(5f);

                //ENCABEZADOS DE TABLA DE INFORMACION
                String[] headers = {"PÁGINA", "FUNCIÓN", "MÓDULO", "APLICACIÓN WEB", "APLICACIÓN MÓVIL"};
                for (String h : headers) {
                    tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                }

                //FILAS CON EFECTO CEBRA
                boolean zebra = false;
                for (FuncionDTO f : rol.getFunciones()) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    zebra = !zebra;

                    tabla.addCell(ReporteUtil.crearCelda(f.getPagina(), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(f.getAccion(), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(transformarModulo(f.getNombre_modulo()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(f.isMovil() ? "" : "Sí", ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(f.isMovil() ? "Sí" : "", ReporteUtil.fuenteTablaData(), fondo));
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

    // METODO AUXILIAR PARA CONVERTIR EL DATO A UN TEXTO MAS AMIGABLE
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
