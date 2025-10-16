package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TimbreDTO {
    private String fechaHora;
    private String accion;
    private String horaTimbre;
}
 