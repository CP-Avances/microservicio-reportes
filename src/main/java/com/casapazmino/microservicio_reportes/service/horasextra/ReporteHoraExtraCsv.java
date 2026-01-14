package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.DataHoraExtra;
import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;

import java.nio.charset.StandardCharsets;

import org.springframework.stereotype.Component;

@Component("reporteHoraExtraCsv")
public class ReporteHoraExtraCsv implements ReporteFile {

    @Override
    public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request) {
        final String EOL = "\r\n";

        try {
            StringBuilder sb = new StringBuilder();

            // Metadata simple (comentarios CSV)
            sb.append("# Empresa: ").append(request.getEmpresa() != null ? request.getEmpresa() : "").append(EOL);
            sb.append("# Usuario: ").append(request.getUsuario() != null ? request.getUsuario() : "").append(EOL);
            sb.append("# Reporte: Horas Extra").append(EOL);

            // Encabezados
            sb.append("fecha,hora,minutos,segundos").append(EOL);

            if (request.getDataHoraExtra() != null) {
                for (DataHoraExtra item : request.getDataHoraExtra()) {
                    String fecha = ReporteUtil.formatearFechaConDia(item.getFecha());
                    sb.append(UtilCsv.csvEscape(fecha)).append(",")
                            .append(UtilCsv.csvEscape(item.getHora())).append(",")
                            .append(UtilCsv.csvEscape(item.getMinutos())).append(",")
                            .append(UtilCsv.csvEscape(item.getSegundos()))
                            .append(EOL);
                }
            }

            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteHorasExtra.csv", e);
        }
    }
}
