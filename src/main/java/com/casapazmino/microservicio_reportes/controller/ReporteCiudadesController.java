package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteCiudadesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteCiudadesService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteCiudadesController {

    @Autowired
    private ReporteCiudadesService reporteService;

    @PostMapping("/ciudades/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteCiudadesRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ciudades.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
