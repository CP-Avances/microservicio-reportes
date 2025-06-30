package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.EstadoCivil.ReporteEstadosCivilRequest;
import com.casapazmino.microservicio_reportes.service.ReporteEstadoCivilService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteEstadoCivilController {

    @Autowired
    private ReporteEstadoCivilService reporteService;

    @PostMapping("/estado-civil/pdf")
    public ResponseEntity<byte[]> generarReporteEstadosCivil(@RequestBody ReporteEstadosCivilRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=estados_civil.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
