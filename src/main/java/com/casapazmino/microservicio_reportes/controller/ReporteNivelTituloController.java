package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.NivelTitulo.ReporteNivelesTitulosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteNivelTituloService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteNivelTituloController {

    @Autowired
    private ReporteNivelTituloService reporteService;

    @PostMapping("/niveles-titulos/pdf")
    public ResponseEntity<byte[]> generarReporteNivelTitulos(@RequestBody ReporteNivelesTitulosRequest request) {
        byte[] pdfBytes = reporteService.generarReporteNivelTituloPDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=niveles_titulos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
