package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteVacunacionUsuariosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteVacunacionUsuariosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/vacunacion-usuarios")
@CrossOrigin(origins = "*")
public class ReporteVacunacionUsuariosController {

    @Autowired
    private ReporteVacunacionUsuariosService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReportePDF(@RequestBody ReporteVacunacionUsuariosRequest request) {
        byte[] pdf = reporteService.generarReportePDF(request);

        if (pdf == null) {
            return ResponseEntity.badRequest().build();
        }

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("filename", "Vacunacion_Usuarios.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdf);
    }
}
