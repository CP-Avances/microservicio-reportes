package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteNacionalidadesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteNacionalidadesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteNacionalidadesController {

    @Autowired
    private ReporteNacionalidadesService reporteService;

    @PostMapping("/nacionalidades/pdf")
    public ResponseEntity<byte[]> generarReporteNacionalidades(@RequestBody ReporteNacionalidadesRequest request) {
        byte[] pdfBytes = reporteService.generarReporteNacionalidadesPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=nacionalidades.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
