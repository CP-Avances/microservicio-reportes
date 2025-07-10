package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CampoTiempoDTO {
    private String fecha_horario;
    private String fecha_hora_horario;
    private String fecha_hora_timbre;
}
