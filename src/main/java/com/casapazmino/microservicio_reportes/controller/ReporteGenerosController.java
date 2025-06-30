package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Genero.ReporteGenerosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteGeneroService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/generos")
public class ReporteGenerosController {

    @Autowired
    private ReporteGeneroService reporteGeneroService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteGeneros(@RequestBody ReporteGenerosRequest request) {
        byte[] pdfBytes = reporteGeneroService.generarReporteGenerosPDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=lista_generos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
