package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones.ReporteKardexVacacionesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteKardexVacacionesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/kardex-vacaciones")
public class ReporteKardexVacacionesController {

    @Autowired
    private ReporteKardexVacacionesService reporteKardexVacacionesService;

    // ===================== PDF =====================
    @PostMapping(
            value = "/pdf",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = MediaType.APPLICATION_PDF_VALUE
    )
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteKardexVacacionesRequest request) {
        try {
            byte[] bin = reporteKardexVacacionesService.generarReporteKardexVacacionesPDF(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=KardexVacaciones.pdf")
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

    // ===================== XLSX =====================
    @PostMapping(
            value = "/xlsx",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    public ResponseEntity<byte[]> generarExcel(@RequestBody ReporteKardexVacacionesRequest request) {
        try {
            byte[] bin = reporteKardexVacacionesService.generarReporteKardexVacacionesExcel(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=KardexVacaciones.xlsx")
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