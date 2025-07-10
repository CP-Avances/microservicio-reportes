package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtualesMovil.ReporteTimbresVirtualesMovilRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTimbresVirtualesMovilService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/timbres-virtuales-movil")
public class ReporteTimbresVirtualesMovilController {

    @Autowired
    private ReporteTimbresVirtualesMovilService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteTimbresVirtualesMovilRequest request) {
        try {
            byte[] pdf = reporteService.generarReportePDF(request);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Timbres_virtuales_movil.pdf")
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(pdf);

        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }
}
