package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteParametrosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/parametros")
public class ReporteParametrosController {

    @Autowired
    private ReporteParametrosService reporteParametrosService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteParametros(@RequestBody ReporteParametrosRequest request) {
        byte[] pdfBytes = reporteParametrosService.generarReporteParametrosPDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_parametros.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteParametrosXlsx(@RequestBody ReporteParametrosRequest request) {
        byte[] xlsxBytes = reporteParametrosService.generarReporteParametrosXLSX(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "ParametrosGeneralesEXCEL.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(xlsxBytes);
    }

    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarReporteParametrosCsv(@RequestBody ReporteParametrosRequest request) {
        byte[] csvBytes = reporteParametrosService.generarReporteParametrosCSV(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("text/csv"));
        headers.setContentDispositionFormData("attachment", "ParametrosGeneralesCSV.csv");
 
        return ResponseEntity.ok()
                .headers(headers)
                .body(csvBytes);
    }

    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarReporteParametrosXml(@RequestBody ReporteParametrosRequest request) {
        byte[] xmlBytes = reporteParametrosService.generarReporteParametrosXML(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_XML);
        headers.setContentDispositionFormData("attachment", "ParametrosGenerales.xml");

        return ResponseEntity.ok()
                .headers(headers)
                .body(xmlBytes);
    }

}
