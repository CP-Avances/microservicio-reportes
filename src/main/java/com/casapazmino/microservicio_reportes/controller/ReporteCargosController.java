package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Cargo.ReporteCargosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteCargosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/cargos")
public class ReporteCargosController {

    @Autowired
    private ReporteCargosService reporteService;

    // ===================== PDF =====================
    @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteCargosRequest request) {
        try {
            byte[] bin = reporteService.generarReportePDF(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Cargos.pdf")
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
    @PostMapping(value = "/xlsx", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteCargosRequest request) {
        try {
            byte[] bin = reporteService.generarReporteXLSX(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType
                            .parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Cargos.xlsx")
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

    // ===================== CSV =====================
    @PostMapping(value = "/csv", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "text/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteCargosRequest request) {
        try {
            byte[] bin = reporteService.generarReporteCSV(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Cargos.csv")
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

    // ===================== XML =====================
    @PostMapping(value = "/xml", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_XML_VALUE)
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteCargosRequest request) {
        try {
            byte[] bin = reporteService.generarReporteXML(request);
            if (bin == null)
                return ResponseEntity.status(500).build();

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_XML)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Cargos.xml")
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
