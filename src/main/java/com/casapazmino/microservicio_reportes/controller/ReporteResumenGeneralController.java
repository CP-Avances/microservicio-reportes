package com.casapazmino.microservicio_reportes.controller;


import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;
import com.casapazmino.microservicio_reportes.service.resumengeneral.interfaces.ReporteResumenGeneralUseCase;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/reporte/resumen-general")
public class ReporteResumenGeneralController {
private final ReporteResumenGeneralUseCase reporteResumenGeneralUseCase;

 public ReporteResumenGeneralController(ReporteResumenGeneralUseCase reporteResumenGeneralUseCase) {
    this.reporteResumenGeneralUseCase = reporteResumenGeneralUseCase;
 }

 @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
 public ResponseEntity<byte[]> generarReporteResumenGeneralPdf(@RequestBody ResumenGeneral request) {
    try {
        byte[] bin = reporteResumenGeneralUseCase.generarPdf(request);
        if (bin == null) return ResponseEntity.status(500).build();
        return ResponseEntity.ok().contentType(MediaType.APPLICATION_PDF).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ResumenGeneral.pdf").header(HttpHeaders.CACHE_CONTROL, "no-store").header(HttpHeaders.PRAGMA, "no-cache").header(HttpHeaders.EXPIRES, "0").body(bin);
    } catch (IllegalArgumentException e) {
        return ResponseEntity.badRequest().build();
    } catch (Exception e) {
        return ResponseEntity.status(500).build();
    }
 }

 @PostMapping(value = "/csv", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "text/csv")
 public ResponseEntity<byte[]> generarReporteResumenGeneralCsv(@RequestBody ResumenGeneral request) {
    try {
        byte[] bin = reporteResumenGeneralUseCase.generarCsv(request);
        if (bin == null) return ResponseEntity.status(500).build();
        return ResponseEntity.ok().contentType(MediaType.parseMediaType("text/csv")).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ResumenGeneral.csv").header(HttpHeaders.CACHE_CONTROL, "no-store").header(HttpHeaders.PRAGMA, "no-cache").header(HttpHeaders.EXPIRES, "0").body(bin);
    } catch (IllegalArgumentException e) {
        return ResponseEntity.badRequest().build();
    } catch (Exception e) {
        return ResponseEntity.status(500).build();
    }
 }

 @PostMapping(value = "/excel", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
 public ResponseEntity<byte[]> generarReporteResumenGeneralExcel(@RequestBody ResumenGeneral request) {
    try {
        byte[] bin = reporteResumenGeneralUseCase.generarExcel(request);
        if (bin == null) return ResponseEntity.status(500).build();
        return ResponseEntity.ok().contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ResumenGeneral.xlsx").header(HttpHeaders.CACHE_CONTROL, "no-store").header(HttpHeaders.PRAGMA, "no-cache").header(HttpHeaders.EXPIRES, "0").body(bin);
    } catch (IllegalArgumentException e) {
        return ResponseEntity.badRequest().build();
    } catch (Exception e) {
        return ResponseEntity.status(500).build();
    }
 }

 @PostMapping(value = "/html", consumes = MediaType.APPLICATION_JSON_VALUE, produces = "text/html")
 public ResponseEntity<byte[]> generarReporteResumenGeneralHtml(@RequestBody ResumenGeneral request) {
    try {
        byte[] bin = reporteResumenGeneralUseCase.generarHtml(request);
        if (bin == null) return ResponseEntity.status(500).build();
        return ResponseEntity.ok().contentType(MediaType.parseMediaType("text/html")).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ResumenGeneral.html").header(HttpHeaders.CACHE_CONTROL, "no-store").header(HttpHeaders.PRAGMA, "no-cache").header(HttpHeaders.EXPIRES, "0").body(bin);
    } catch (IllegalArgumentException e) {
        return ResponseEntity.badRequest().build();
    } catch (Exception e) {
        return ResponseEntity.status(500).build();
    }
 }

}
