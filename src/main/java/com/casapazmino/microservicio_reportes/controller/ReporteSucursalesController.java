package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteSucursalesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteSucursalesService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteSucursalesController {

    @Autowired
    private ReporteSucursalesService reporteService;

    @PostMapping("/sucursales/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteSucursalesRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sucursales.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
