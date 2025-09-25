package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Genero.ReporteGenerosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteGeneroService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/generos")
@CrossOrigin(origins = "http://localhost:51871") //TEMA A INVESTIGAR
public class ReporteGenerosController {

    @Autowired
    private ReporteGeneroService reporteGeneroService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteGeneros(@RequestBody ReporteGenerosRequest request) {
        byte[] pdfBytes = reporteGeneroService.generarReporteGenerosPDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=lista_generos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

        // XLSX
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteGenerosRequest request) {
        byte[] bin = reporteGeneroService.generarReporteGenerosXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=generos.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteGenerosRequest request) {
        byte[] bin = reporteGeneroService.generarReporteGenerosCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=generos.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteGenerosRequest request) {
        byte[] bin = reporteGeneroService.generarReporteGenerosXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=generos.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
 