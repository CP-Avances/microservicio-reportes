package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Vacuna.ReporteVacunasRequest;
import com.casapazmino.microservicio_reportes.service.ReporteVacunaService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/vacunas")
public class ReporteVacunaController {

    private final ReporteVacunaService reporteService;

    public ReporteVacunaController(ReporteVacunaService reporteService) {
        this.reporteService = reporteService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteVacunas(@RequestBody ReporteVacunasRequest request) {
        byte[] pdf = reporteService.generarReporteVacunasPDF(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Vacunas.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
