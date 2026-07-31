package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio.ReporteTiempoServicioRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTiempoServicioService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/tiempo-servicio")
public class ReporteTiempoServicioController {

    @Autowired
    private ReporteTiempoServicioService reporteService;

    // ===================== PDF =====================

    @PostMapping(
            value = "/pdf",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = MediaType.APPLICATION_PDF_VALUE
    )
    public ResponseEntity<byte[]> generarPDF(
            @RequestBody ReporteTiempoServicioRequest request
    ) {
        try {
            byte[] bin =
                    reporteService.generarReporteTiempoServicioPDF(request);

            if (bin == null || bin.length == 0) {
                return ResponseEntity
                        .status(500)
                        .build();
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(
                            HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=TiempoServicio.pdf"
                    )
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);

        } catch (IllegalArgumentException e) {
            return ResponseEntity
                    .badRequest()
                    .build();

        } catch (Exception e) {
            return ResponseEntity
                    .status(500)
                    .build();
        }
    }

    // ===================== XLSX =====================

    @PostMapping(
            value = "/xlsx",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    public ResponseEntity<byte[]> generarExcel(
            @RequestBody ReporteTiempoServicioRequest request
    ) {
        try {
            byte[] bin =
                    reporteService.generarReporteTiempoServicioExcel(request);

            if (bin == null || bin.length == 0) {
                return ResponseEntity
                        .status(500)
                        .build();
            }

            MediaType excelMediaType =
                    MediaType.parseMediaType(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    );

            return ResponseEntity.ok()
                    .contentType(excelMediaType)
                    .header(
                            HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=TiempoServicio.xlsx"
                    )
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);

        } catch (IllegalArgumentException e) {
            return ResponseEntity
                    .badRequest()
                    .build();

        } catch (Exception e) {
            return ResponseEntity
                    .status(500)
                    .build();
        }
    }
}