package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteParametrosService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/parametros")
public class ReporteParametrosController {

    @Autowired
    private ReporteParametrosService reporteParametrosService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteParametros(@RequestBody ReporteParametrosRequest request) {
        byte[] pdfBytes = reporteParametrosService.generarReporteParametrosPDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_parametros.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }
}
