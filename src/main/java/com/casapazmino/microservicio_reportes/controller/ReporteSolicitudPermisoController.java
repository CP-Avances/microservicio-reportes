package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.SolicitudPermiso.ReporteSolicitudPermisoRequest;
import com.casapazmino.microservicio_reportes.service.ReporteSolicitudPermisoService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/solicitud-permiso")
public class ReporteSolicitudPermisoController {

    @Autowired
    private ReporteSolicitudPermisoService reporteSolicitudPermisoService;

    // ===================== PDF =====================
    @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
    public ResponseEntity<byte[]> generarReporteSolicitudPermiso(@RequestBody ReporteSolicitudPermisoRequest request) {
        try {
            byte[] bin = reporteSolicitudPermisoService.generarReporteSolicitudPermisoPDF(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=SolicitudPermiso.pdf")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build();
        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }
}
