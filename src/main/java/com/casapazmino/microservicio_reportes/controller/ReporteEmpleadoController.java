package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteEmpleadosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteEmpleadoService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/empleados")
public class ReporteEmpleadoController {

    private final ReporteEmpleadoService reporteEmpleadoService;

    public ReporteEmpleadoController(ReporteEmpleadoService reporteEmpleadoService) {
        this.reporteEmpleadoService = reporteEmpleadoService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteEmpleado(@RequestBody ReporteEmpleadosRequest request) {
        byte[] pdf = reporteEmpleadoService.generarReporteEmpleadosPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Empleados.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}
