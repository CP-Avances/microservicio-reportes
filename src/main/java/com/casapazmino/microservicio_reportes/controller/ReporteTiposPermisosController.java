package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.TipoPermiso.ReporteTiposPermisosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTiposPermisosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/tipos-permisos")
public class ReporteTiposPermisosController {

    @Autowired
    private ReporteTiposPermisosService reporteTiposPermisosService;

    // ===================== PDF =====================
    @PostMapping(
            value = "/pdf",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = MediaType.APPLICATION_PDF_VALUE
    )
    public ResponseEntity<byte[]> generarReporteTiposPermisosPdf(
            @RequestBody ReporteTiposPermisosRequest request) {
        try {
            byte[] bin = reporteTiposPermisosService.generarReporteTiposPermisosPDF(request);

            if (bin == null) {
                return ResponseEntity.status(500).build();
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(
                            HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=TiposPermisos.pdf"
                    )
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
    @PostMapping(
            value = "/xlsx",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    public ResponseEntity<byte[]> generarReporteTiposPermisosXlsx(
            @RequestBody ReporteTiposPermisosRequest request) {
        try {
            byte[] bin = reporteTiposPermisosService.generarReporteTiposPermisosXLSX(request);

            if (bin == null) {
                return ResponseEntity.status(500).build();
            }

            return ResponseEntity.ok()
                    .contentType(
                            MediaType.parseMediaType(
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    )
                    .header(
                            HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=TiposPermisos.xlsx"
                    )
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
    @PostMapping(
            value = "/csv",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = "text/csv"
    )
    public ResponseEntity<byte[]> generarReporteTiposPermisosCsv(
            @RequestBody ReporteTiposPermisosRequest request) {
        try {
            byte[] bin = reporteTiposPermisosService.generarReporteTiposPermisosCSV(request);

            if (bin == null) {
                return ResponseEntity.status(500).build();
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
                    .header(
                            HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=TiposPermisos.csv"
                    )
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
    @PostMapping(
            value = "/xml",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = MediaType.APPLICATION_XML_VALUE
    )
    public ResponseEntity<byte[]> generarReporteTiposPermisosXml(
            @RequestBody ReporteTiposPermisosRequest request) {
        try {
            byte[] bin = reporteTiposPermisosService.generarReporteTiposPermisosXML(request);

            if (bin == null) {
                return ResponseEntity.status(500).build();
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_XML)
                    .header(
                            HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=TiposPermisos.xml"
                    )
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