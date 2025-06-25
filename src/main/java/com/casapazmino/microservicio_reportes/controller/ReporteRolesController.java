package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteRolesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteRolesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteRolesController {

    @Autowired
    private ReporteRolesService reporteRolesService;

    @PostMapping("/roles/pdf")
    public ResponseEntity<byte[]> generarReporteRoles(@RequestBody ReporteRolesRequest request) {
        byte[] pdfBytes = reporteRolesService.generarReporteRolesPDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=lista_roles.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
