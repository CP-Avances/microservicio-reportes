package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresUsuarios.ReporteTimbresUsuariosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTimbresUsuariosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

@RestController
@RequestMapping("/api/reportes/timbres-usuarios")
public class ReporteTimbresUsuariosController {

    @Autowired
    private ReporteTimbresUsuariosService reporteService;

    @PostMapping(
        path = "/pdf",
        produces = MediaType.APPLICATION_PDF_VALUE
    )
    public ResponseEntity<StreamingResponseBody> generarPDF(@RequestBody ReporteTimbresUsuariosRequest request) {

        StreamingResponseBody stream = outputStream -> {
            // El service debe **escribir** directo al OutputStream
            try {
                reporteService.escribirReportePDF(request, outputStream);
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            // Importante: NO cerrar outputStream aquí; Spring lo maneja.
        };

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Timbres_usuarios.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(stream);
    }

    @PostMapping(
        path = "/xlsx",
        produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    public ResponseEntity<StreamingResponseBody> generarExcel(@RequestBody ReporteTimbresUsuariosRequest request) {

        StreamingResponseBody stream = outputStream -> {
            // El service debe **escribir** directo al OutputStream
            try {
                reporteService.escribirReporteTimbresUsuariosExcel(request, outputStream);
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        };

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=timbres_usuarios.xlsx")
                .header(HttpHeaders.CACHE_CONTROL, "no-store")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(stream);
    }
}
