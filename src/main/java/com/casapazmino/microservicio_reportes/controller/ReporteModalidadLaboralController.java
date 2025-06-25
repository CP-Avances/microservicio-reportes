package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteModalidadLaboralRequest;
import com.casapazmino.microservicio_reportes.service.ReporteModalidadLaboralService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteModalidadLaboralController {

    @Autowired
    private ReporteModalidadLaboralService reporteService;

    @PostMapping("/modalidad-laboral/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteModalidadLaboralRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=modalidad_laboral.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
