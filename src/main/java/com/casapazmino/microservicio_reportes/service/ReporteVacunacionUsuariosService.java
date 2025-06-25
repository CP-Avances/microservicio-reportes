package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;

@Service
public class ReporteVacunacionUsuariosService {

    public byte[] generarReportePDF(ReporteVacunacionUsuariosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4, 40, 40, 50, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                logo.scaleAbsoluteWidth(100);
                logo.setAlignment(Image.ALIGN_LEFT);
                document.add(logo);
            }

            Paragraph empresa = new Paragraph(request.getEmpresa(), ReporteUtil.fuenteEncabezado());
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            Paragraph titulo = new Paragraph(safe(request.getTitulo()), ReporteUtil.fuenteEncabezado());
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Font fuente = ReporteUtil.fuenteTexto();

            for (AgrupadorVacunaUsuarioDTO grupo : request.getDatos()) {

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

                int totalRegistros = grupo.getEmpleados().stream()
                        .mapToInt(e -> e.getVacunas().size())
                        .sum();

                // Tabla verde sin bordes internos (solo borde externo)
                PdfPTable tablaCabecera = new PdfPTable(3);
                tablaCabecera.setWidthPercentage(100);
                tablaCabecera.setWidths(new float[]{3, 3, 2});
                tablaCabecera.setSpacingBefore(10f);
                tablaCabecera.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                tablaCabecera.addCell(celdaSinBordeIzquierda(descripcion, fuente, colorSecundario));
                tablaCabecera.addCell(celdaSinBordeIzquierda(establecimiento, fuente, colorSecundario));
                tablaCabecera.addCell(celdaSinBordeIzquierda("N° Registros: " + totalRegistros, fuente, colorSecundario));

                tablaCabecera.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                    PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                    cb.rectangle(
                            widths[0][0],
                            heights[heights.length - 1],
                            widths[0][widths[0].length - 1] - widths[0][0],
                            heights[0] - heights[heights.length - 1]
                    );
                    cb.stroke();
                });

                document.add(tablaCabecera);

                for (EmpleadoVacunaUsuarioDTO empl : grupo.getEmpleados()) {
                    PdfPTable tablaEmpleado = new PdfPTable(3);
                    tablaEmpleado.setWidthPercentage(100);
                    tablaEmpleado.setSpacingBefore(6f);
                    tablaEmpleado.setWidths(new float[]{3, 4, 3});

                    String[][] filas = new String[][]{
                            {
                                    "C.C.: " + safe(empl.getIdentificacion()),
                                    "EMPLEADO: " + safe(empl.getApellido()) + " " + safe(empl.getNombre()),
                                    "DEPARTAMENTO: " + safe(empl.getDepartamento())
                            },
                            {
                                    "CORREO: " + safe(empl.getCorreo()),
                                    "GENERO: " + safe(empl.getGenero()),
                                    "CARGO: " + safe(empl.getCargo())
                            },
                            {
                                    "REGIMEN: " + safe(empl.getRegimen()),
                                    "COD: " + safe(empl.getCodigo()),
                                    "ROL: " + safe(empl.getRol())
                            }
                    };

                    for (String[] fila : filas) {
                        for (String texto : fila) {
                            PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
                            celda.setBackgroundColor(Color.LIGHT_GRAY);
                            celda.setHorizontalAlignment(Element.ALIGN_LEFT);
                            celda.setBorder(Rectangle.NO_BORDER);
                            tablaEmpleado.addCell(celda);
                        }
                    }

                    tablaEmpleado.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                        PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                        cb.rectangle(
                                widths[0][0],
                                heights[heights.length - 1],
                                widths[0][widths[0].length - 1] - widths[0][0],
                                heights[0] - heights[heights.length - 1]
                        );
                        cb.stroke();
                    });

                    document.add(tablaEmpleado);

                    PdfPTable tablaVacunas = new PdfPTable(4);
                    tablaVacunas.setWidthPercentage(100);
                    tablaVacunas.setSpacingBefore(0f);
                    tablaVacunas.setWidths(new float[]{1, 3, 2, 5});

                    // Encabezados centrados
                    for (String header : new String[]{"N°", "VACUNA", "FECHA", "DESCRIPCIÓN"}) {
                        PdfPCell headerCell = new PdfPCell(new Phrase(header, fuente));
                        headerCell.setBackgroundColor(colorPrincipal);
                        headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        tablaVacunas.addCell(headerCell);
                    }

                    int index = 1;
                    for (VacunaUsuarioDTO vac : empl.getVacunas()) {
                        PdfPCell celda1 = new PdfPCell(new Phrase(String.valueOf(index++), fuente));
                        PdfPCell celda2 = new PdfPCell(new Phrase(safe(vac.getTipo_vacuna()), fuente));
                        PdfPCell celda3 = new PdfPCell(new Phrase(formatearFecha(vac.getFecha()), fuente));
                        PdfPCell celda4 = new PdfPCell(new Phrase(safe(vac.getDescripcion()), fuente));

                        for (PdfPCell celda : new PdfPCell[]{celda1, celda2, celda3, celda4}) {
                            celda.setHorizontalAlignment(Element.ALIGN_CENTER);
                        }

                        tablaVacunas.addCell(celda1);
                        tablaVacunas.addCell(celda2);
                        tablaVacunas.addCell(celda3);
                        tablaVacunas.addCell(celda4);
                    }

                    document.add(tablaVacunas);
                }
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell celdaSinBordeIzquierda(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setBorder(Rectangle.NO_BORDER);
        celda.setPaddingTop(6f);
        celda.setPaddingBottom(6f);
        return celda;
    }

    private String safe(String val) {
        return (val == null || val.equalsIgnoreCase("null")) ? "" : val;
    }

    private String formatearFecha(String fechaIso) {
        try {
            LocalDate fecha = LocalDate.parse(fechaIso.substring(0, 10));
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEE. dd/MM/yyyy", new Locale("es", "ES"));
            return formatter.format(fecha);
        } catch (Exception e) {
            return fechaIso;
        }
    }
}
