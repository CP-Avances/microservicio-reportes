package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteUsuario.ReporteUsuariosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteUsuariosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteUsuariosController {

    @Autowired
    private ReporteUsuariosService reporteService;

    @PostMapping("/usuarios/pdf")
    public ResponseEntity<byte[]> generarReporteUsuarios(@RequestBody ReporteUsuariosRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=usuarios.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    // === XLSX ===
    @PostMapping("/usuarios/xlsx")
    public ResponseEntity<byte[]> generarReporteUsuariosXLSX(@RequestBody ReporteUsuariosRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Usuarios.xlsx")
                .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }
}
