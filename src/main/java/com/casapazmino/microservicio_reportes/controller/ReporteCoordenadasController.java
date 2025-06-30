package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Coordenada.ReporteCoordenadasRequest;
import com.casapazmino.microservicio_reportes.service.ReporteCoordenadasService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/coordenadas")
public class ReporteCoordenadasController {

    @Autowired
    private ReporteCoordenadasService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteCoordenadasRequest request) {
        byte[] pdf = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=CoordenadasGeograficas.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
