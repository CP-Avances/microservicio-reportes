package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ResumenAsistencia.ReporteAsistenciaRequest;
import com.casapazmino.microservicio_reportes.service.ReporteAsistenciaService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/reporte/asistencia")
public class ReporteAsistenciaController {

    @Autowired
    private ReporteAsistenciaService reporteAsistenciaService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteAsistencia(@RequestBody ReporteAsistenciaRequest request) {
        byte[] pdfBytes = reporteAsistenciaService.generarReporteResumenAsistenciaPDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_asistencia.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }
}