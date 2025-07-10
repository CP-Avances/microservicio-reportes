package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class SalidaDTO {
    private String fecha_hora_horario;
    private String fecha_hora_timbre;
    private Double diferencia;
    private String tipo_permiso;
    private String desde;
    private String hasta;

}
