package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteAuditoria.ReporteAuditoriaRequest;
import com.casapazmino.microservicio_reportes.service.ReporteAuditoriaService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/auditoria")
public class ReporteAuditoriaController {

    @Autowired
    private ReporteAuditoriaService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteAuditoriaRequest request) {
        try {
            byte[] pdf = reporteService.generarReportePDF(request);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Auditoria.pdf")
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(pdf);

        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }
}
