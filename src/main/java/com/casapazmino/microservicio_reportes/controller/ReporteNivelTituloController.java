package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.NivelTitulo.ReporteNivelesTitulosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteNivelTituloService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteNivelTituloController {

    @Autowired
    private ReporteNivelTituloService reporteService;

    @PostMapping("/niveles-titulos/pdf")
    public ResponseEntity<byte[]> generarReporteNivelTitulos(@RequestBody ReporteNivelesTitulosRequest request) {
        byte[] pdfBytes = reporteService.generarReporteNivelTituloPDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=niveles_titulos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    // XLSX
    @PostMapping("/niveles-titulos/xlsx")
    public ResponseEntity<byte[]> generarReporteNivelTitulosXLSX(@RequestBody ReporteNivelesTitulosRequest request) {
        byte[] bin = reporteService.generarReporteNivelTituloXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=niveles_titulos.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/niveles-titulos/csv")
    public ResponseEntity<byte[]> generarReporteNivelTitulosCSV(@RequestBody ReporteNivelesTitulosRequest request) {
        byte[] bin = reporteService.generarReporteNivelTituloCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=niveles_titulos.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/niveles-titulos/xml")
    public ResponseEntity<byte[]> generarReporteNivelTitulosXML(@RequestBody ReporteNivelesTitulosRequest request) {
        byte[] bin = reporteService.generarReporteNivelTituloXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=niveles_titulos.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
