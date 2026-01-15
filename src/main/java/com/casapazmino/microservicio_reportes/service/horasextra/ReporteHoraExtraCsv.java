package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import java.nio.charset.StandardCharsets;

import org.springframework.stereotype.Component;

@Component("reporteHoraExtraCsv")
public class ReporteHoraExtraCsv implements ReporteFile {

    @Override
    public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request) {
        final String EOL = "\r\n";

        try {
            StringBuilder stringBuilder = new StringBuilder();

            // Metadata simple (comentarios CSV)
            stringBuilder.append("# Empresa: ").append(request.getEmpresa() != null ? request.getEmpresa() : "").append(EOL);
            stringBuilder.append("# Usuario: ").append(request.getUsuario() != null ? request.getUsuario() : "").append(EOL);
            stringBuilder.append("# Reporte: Horas Extra").append(EOL);

            // Encabezados
            stringBuilder.append("fecha,hora,minutos,segundos").append(EOL);



            return stringBuilder.toString().getBytes(StandardCharsets.UTF_8);
        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteHorasExtra.csv", e);
        }
    }
}
