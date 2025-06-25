package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteFeriadosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteFeriadosService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/feriados")
public class ReporteFeriadoController {

    private final ReporteFeriadosService reporteFeriadosService;

    public ReporteFeriadoController(ReporteFeriadosService reporteFeriadosService) {
        this.reporteFeriadosService = reporteFeriadosService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteFeriadosRequest request) {
        byte[] pdf = reporteFeriadosService.generarReporteFeriadosPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=feriados.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
