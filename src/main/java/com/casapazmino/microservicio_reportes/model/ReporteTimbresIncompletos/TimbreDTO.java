package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TimbreDTO {
    private String fechaHora;
    private String accion;
    private String horaTimbre;
}
