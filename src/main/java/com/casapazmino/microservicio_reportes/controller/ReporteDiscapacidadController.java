package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Discapacidad.ReporteDiscapacidadesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteDiscapacidadService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/discapacidades")
@CrossOrigin(origins = "*")
public class ReporteDiscapacidadController {

    private final ReporteDiscapacidadService reporteService;

    public ReporteDiscapacidadController(ReporteDiscapacidadService reporteService) {
        this.reporteService = reporteService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteDiscapacidades(@RequestBody ReporteDiscapacidadesRequest request) {
        byte[] pdf = reporteService.generarReporteDiscapacidadesPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Discapacidades.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
