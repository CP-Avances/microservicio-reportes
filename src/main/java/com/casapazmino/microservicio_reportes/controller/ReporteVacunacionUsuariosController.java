package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteVacunacion.ReporteVacunacionUsuariosRequest;
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
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("filename", "Vacunacion_Usuarios.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdf);
    }

    // XLSX
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteVacunacionUsuariosRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Vacunacion_Usuarios.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }
}
