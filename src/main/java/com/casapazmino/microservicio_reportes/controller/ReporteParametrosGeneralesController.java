package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosGeneralesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteParametrosGeneralesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/parametros-generales")
public class ReporteParametrosGeneralesController {

    @Autowired
    private ReporteParametrosGeneralesService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPdf(@RequestBody ReporteParametrosGeneralesRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Parametros_Generales.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
