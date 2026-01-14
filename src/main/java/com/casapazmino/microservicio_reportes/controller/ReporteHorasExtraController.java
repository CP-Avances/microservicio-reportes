package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.service.horasextra.ReporteHorasExtraUseCase;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/reporte/horas-extra")
public class ReporteHorasExtraController {
    @Autowired
    private ReporteHorasExtraUseCase reporteHorasExtraService;

    @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
    public ResponseEntity<byte[]> generarReporteHorasExtra(@RequestBody ReporteHorasExtraRequest request) {
        try {
            byte[] bin = reporteHorasExtraService.generarPdf(request);
            if (bin == null) return ResponseEntity.status(500).build();

            return ResponseEntity.ok().contentType(MediaType.APPLICATION_PDF).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=HorasExtra.pdf").header(HttpHeaders.CACHE_CONTROL, "no-store").header(HttpHeaders.PRAGMA, "no-cache").header(HttpHeaders.EXPIRES, "0").body(bin);
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build();
        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }

    @PostMapping(value = "/csv", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "text/csv")
    public ResponseEntity<byte[]> generarReporteHorasExtraCsv(@RequestBody ReporteHorasExtraRequest request) {
        try {
            byte[] bin = reporteHorasExtraService.generarCsv(request);
            if (bin == null) return ResponseEntity.status(500).build();

            return ResponseEntity.ok().contentType(MediaType.parseMediaType("text/csv")).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=HorasExtra.csv").header(HttpHeaders.CACHE_CONTROL, "no-store").header(HttpHeaders.PRAGMA, "no-cache").header(HttpHeaders.EXPIRES, "0").body(bin);
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().build();
        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }

}
