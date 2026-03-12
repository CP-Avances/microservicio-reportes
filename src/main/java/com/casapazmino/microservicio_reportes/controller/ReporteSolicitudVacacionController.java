package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.SolicitudVacacion.ReporteSolicitudVacacionRequest;
import com.casapazmino.microservicio_reportes.service.ReporteSolicitudVacacionService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/solicitud-vacacion")
public class ReporteSolicitudVacacionController {

    @Autowired
    private ReporteSolicitudVacacionService reporteSolicitudVacacionService;

    @PostMapping(
            value = "/pdf",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = MediaType.APPLICATION_PDF_VALUE
    )
    public ResponseEntity<byte[]> generarReporteSolicitudVacacion(@RequestBody ReporteSolicitudVacacionRequest request) {
        try {
            byte[] bin = reporteSolicitudVacacionService.generarReporteSolicitudVacacionPDF(request);
            if (bin == null) return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=SolicitudVacacion.pdf")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build();
        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }
}