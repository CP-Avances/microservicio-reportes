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
@RequestMapping("/api/reporte/timbres-usuarios")
public class ReporteTimbresUsuariosController {

    @Autowired
    private ReporteTimbresUsuariosService reporteService;

    // ===================== PDF (Streaming) =====================
    @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
    public ResponseEntity<StreamingResponseBody> generarPDF(@RequestBody ReporteTimbresUsuariosRequest request) {
        try {
            StreamingResponseBody stream = outputStream -> {
                try {
                    // El service escribe directo al OutputStream (sin cerrar)
                    reporteService.escribirReportePDF(request, outputStream);
                } catch (IllegalArgumentException iae) {
                    // Propagamos para que Spring responda 400 si aplica
                    throw iae;
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            };

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=TimbresUsuarios.pdf")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(stream);

        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build(); // 400
        } catch (Exception e) {
            return ResponseEntity.status(500).build(); // 500
        }
    }

    // ===================== XLSX (Streaming) =====================
    @PostMapping(value = "/xlsx", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<StreamingResponseBody> generarExcel(@RequestBody ReporteTimbresUsuariosRequest request) {
        try {
            StreamingResponseBody stream = outputStream -> {
                try {
                    // El service escribe directo al OutputStream (sin cerrar)
                    reporteService.escribirReporteTimbresUsuariosExcel(request, outputStream);
                } catch (IllegalArgumentException iae) {
                    throw iae;
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            };

            return ResponseEntity.ok()
                    .contentType(MediaType
                            .parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=TimbresUsuarios.xlsx")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(stream);

        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build(); // 400
        } catch (Exception e) {
            return ResponseEntity.status(500).build(); // 500
        }
    }
}
