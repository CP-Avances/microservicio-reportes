package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Departamento.ReporteDepartamentosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteDepartamentosService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/departamentos")
@CrossOrigin(origins = "*")
public class ReporteDepartamentosController {

    private final ReporteDepartamentosService reporteService;

    public ReporteDepartamentosController(ReporteDepartamentosService reporteService) {
        this.reporteService = reporteService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReportePDF(@RequestBody ReporteDepartamentosRequest request) {
        byte[] pdf = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Departamentos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
