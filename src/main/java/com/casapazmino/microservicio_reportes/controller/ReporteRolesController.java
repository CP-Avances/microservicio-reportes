package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Rol.ReporteRolesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteRolesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteRolesController {

    @Autowired
    private ReporteRolesService reporteRolesService;

    @PostMapping("/roles/pdf")
    public ResponseEntity<byte[]> generarReporteRoles(@RequestBody ReporteRolesRequest request) {
        // El servicio genera el pdf
        byte[] pdfBytes = reporteRolesService.generarReporteRolesPDF(request);

        // Tipo de retorno es un pdf en bytes
        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=lista_roles.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    @PostMapping("/roles/xlsx")
    public ResponseEntity<byte[]> generarReporteRolesXlsx(@RequestBody ReporteRolesRequest request) {
        byte[] xlsxBytes = reporteRolesService.generarReporteRolesXLSX(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "RolesEXCEL.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(xlsxBytes);
    }

    @PostMapping("/roles/csv")
    public ResponseEntity<byte[]> generarReporteRolesCsv(@RequestBody ReporteRolesRequest request) {
        byte[] csvBytes = reporteRolesService.generarReporteRolesCSV(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("text/csv"));
        headers.setContentDispositionFormData("attachment", "RolesCSV.csv");

        return ResponseEntity.ok().headers(headers).body(csvBytes);
    }

    @PostMapping("/roles/xml")
    public ResponseEntity<byte[]> generarReporteRolesXml(@RequestBody ReporteRolesRequest request) {
        byte[] xmlBytes = reporteRolesService.generarReporteRolesXML(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_XML);
        headers.setContentDispositionFormData("attachment", "Roles.xml");

        return ResponseEntity.ok().headers(headers).body(xmlBytes);
    }

}
