package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales.ReporteTimbresVirtualesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTimbresVirtualesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/timbres-virtuales")
public class ReporteTimbresVirtualesController {

    @Autowired
    private ReporteTimbresVirtualesService reporteService;

    // ===================== PDF =====================
    @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteTimbresVirtualesRequest request) {
        try {
            byte[] bin = reporteService.generarReportePDF(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=TimbresVirtuales.pdf")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);

        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build(); // 400
        } catch (Exception e) {
            return ResponseEntity.status(500).build(); // 500
        }
    }

    // ===================== XLSX =====================
    @PostMapping(value = "/xlsx", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<byte[]> generarXLSX(@RequestBody ReporteTimbresVirtualesRequest request) {
        try {
            byte[] bin = reporteService.generarReporteXLSX(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType
                            .parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=TimbresVirtuales.xlsx")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);

        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build(); // 400
        } catch (Exception e) {
            return ResponseEntity.status(500).build(); // 500
        }
    }
}
