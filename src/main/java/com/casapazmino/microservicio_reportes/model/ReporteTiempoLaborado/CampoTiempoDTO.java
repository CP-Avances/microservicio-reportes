package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class CampoTiempoDTO {
    private String fecha_horario;
    private String fecha_hora_horario;
    private String fecha_hora_timbre;
}
