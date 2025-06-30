package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Provincia.ReporteProvinciasRequest;
import com.casapazmino.microservicio_reportes.service.ReporteProvinciasService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteProvinciasController {

    @Autowired
    private ReporteProvinciasService reporteService;

    @PostMapping("/provincias/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteProvinciasRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=provincias.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
