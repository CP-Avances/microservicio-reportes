package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteUsuario.AgrupadorUsuariosDTO;
import com.casapazmino.microservicio_reportes.model.ReporteUsuario.ReporteUsuariosRequest;
import com.casapazmino.microservicio_reportes.model.ReporteUsuario.UsuarioDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteUsuariosService {

    public byte[] generarReportePDF(ReporteUsuariosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 60, 40);
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            Paragraph empresa = new Paragraph(request.getEmpresa(), ReporteUtil.fuenteEncabezado());
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            Paragraph titulo = new Paragraph(request.getTitulo(), ReporteUtil.fuenteEncabezado());
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Font fuente = ReporteUtil.fuenteTexto();

            for (AgrupadorUsuariosDTO grupo : request.getDatos()) {
                String descripcion = "";
                String establecimiento = safe("SUCURSAL: " + grupo.getSucursal());

                switch (safe(request.getTipoFiltro()).toLowerCase()) {
                    case "regimen":
                        descripcion = "REGIMEN: " + safe(grupo.getNombre());
                        break;
                    case "departamento":
                        descripcion = "DEPARTAMENTO: " + safe(grupo.getDepartamento());
                        break;
                    case "cargo":
                        descripcion = "CARGO: " + safe(grupo.getNombre());
                        break;
                    case "ciudad":
                        descripcion = "CIUDAD: " + safe(grupo.getCiudad());
                        break;
                    case "empleado":
                        descripcion = "LISTA EMPLEADOS";
                        establecimiento = "";
                        break;
                }

                PdfPTable cabecera = new PdfPTable(3);
                cabecera.setWidthPercentage(100);
                cabecera.setWidths(new float[]{3, 3, 2});
                cabecera.setSpacingBefore(10f);
                cabecera.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                cabecera.addCell(celdaSinBorde(descripcion, fuente, colorSecundario));
                cabecera.addCell(celdaSinBorde(establecimiento, fuente, colorSecundario));
                cabecera.addCell(celdaSinBorde("N° Registros: " + grupo.getEmpleados().size(), fuente, colorSecundario));

                cabecera.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                    PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                    cb.rectangle(
                            widths[0][0],
                            heights[heights.length - 1],
                            widths[0][widths[0].length - 1] - widths[0][0],
                            heights[0] - heights[heights.length - 1]
                    );
                    cb.stroke();
                });

                document.add(cabecera);

                // Tabla de usuarios
                PdfPTable tablaUsuarios = new PdfPTable(14);
                tablaUsuarios.setWidthPercentage(100);
                tablaUsuarios.setSpacingBefore(0f);
                tablaUsuarios.setWidths(new float[]{1, 3, 3, 4, 3, 2, 3, 3, 3, 3, 3, 3, 3, 5});

                tablaUsuarios.addCell(ReporteUtil.crearCelda("N°", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("IDENTIFICACIÓN", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CÓDIGO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("EMPLEADO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("USUARIO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("GÉNERO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("NACIONALIDAD", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CIUDAD", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("SUCURSAL", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("RÉGIMEN", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("DEPARTAMENTO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CARGO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("ROL", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CORREO", fuente, colorPrincipal));

                int index = 1;
                for (UsuarioDTO usu : grupo.getEmpleados()) {
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(String.valueOf(index++), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getIdentificacion()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCodigo()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaIzquierda(safe(usu.getApellido()) + " " + safe(usu.getNombre()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getUsuario()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getGenero()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getNacionalidad()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCiudad()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getSucursal()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getRegimen()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getDepartamento()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCargo()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getRol()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaIzquierda(safe(usu.getCorreo()), fuente));
                }

                document.add(tablaUsuarios);
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell celdaSinBorde(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setBorder(Rectangle.NO_BORDER);
        celda.setPaddingTop(6f);
        celda.setPaddingBottom(6f);
        return celda;
    }

    private String safe(String val) {
        return (val == null || val.equalsIgnoreCase("null")) ? "" : val;
    }
}
