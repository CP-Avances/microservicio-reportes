package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Regimen.ReporteRegimenRequest;
import com.casapazmino.microservicio_reportes.service.ReporteRegimenService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/regimen")
@CrossOrigin(origins = "*")
public class ReporteRegimenController {

    private final ReporteRegimenService reporteRegimenService;

    public ReporteRegimenController(ReporteRegimenService reporteRegimenService) {
        this.reporteRegimenService = reporteRegimenService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReportePDF(@RequestBody ReporteRegimenRequest request) throws Exception {
        byte[] pdf = reporteRegimenService.generarReporteRegimenPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Regimen_laboral.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
