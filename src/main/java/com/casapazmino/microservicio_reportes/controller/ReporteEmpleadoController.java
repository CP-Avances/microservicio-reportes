package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Empleado.ReporteEmpleadosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteEmpleadoService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/empleados")
public class ReporteEmpleadoController {

    private final ReporteEmpleadoService reporteEmpleadoService;

    public ReporteEmpleadoController(ReporteEmpleadoService reporteEmpleadoService) {
        this.reporteEmpleadoService = reporteEmpleadoService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteEmpleado(@RequestBody ReporteEmpleadosRequest request) {
        byte[] pdf = reporteEmpleadoService.generarReporteEmpleadosPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Empleados.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }

        // XLSX
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteEmpleadoXLSX(@RequestBody ReporteEmpleadosRequest request) {
        byte[] bin = reporteEmpleadoService.generarReporteEmpleadosXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Empleados.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarReporteEmpleadoCSV(@RequestBody ReporteEmpleadosRequest request) {
        byte[] bin = reporteEmpleadoService.generarReporteEmpleadosCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Empleados.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarReporteEmpleadoXML(@RequestBody ReporteEmpleadosRequest request) {
        byte[] bin = reporteEmpleadoService.generarReporteEmpleadosXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Empleados.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
