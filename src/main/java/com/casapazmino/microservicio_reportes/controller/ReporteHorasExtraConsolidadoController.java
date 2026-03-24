package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado.ReporteHorasExtraConsolidadoRequest;
import com.casapazmino.microservicio_reportes.service.ReporteHorasExtraConsolidadoService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/horas-extra-consolidado")
public class ReporteHorasExtraConsolidadoController {

    @Autowired
    private ReporteHorasExtraConsolidadoService reporteHorasExtraConsolidadoService;

    @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
    public ResponseEntity<byte[]> generarReportePDF(
            @RequestBody ReporteHorasExtraConsolidadoRequest request) {
        try {
            byte[] bin = reporteHorasExtraConsolidadoService.generarReportePDF(request);
            if (bin == null) {
                return ResponseEntity.status(500).build();
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=HorasExtraConsolidado.pdf")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);

        } catch (IllegalArgumentException e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().build();
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).build();
        }
    }

    @PostMapping(value = "/xlsx", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<byte[]> generarReporteXLSX(
            @RequestBody ReporteHorasExtraConsolidadoRequest request) {
        try {
            byte[] bin = reporteHorasExtraConsolidadoService.generarReporteXLSX(request);
            if (bin == null) {
                return ResponseEntity.status(500).build();
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=HorasExtraConsolidado.xlsx")
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .body(bin);

        } catch (IllegalArgumentException e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().build();
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).build();
        }
    }
}