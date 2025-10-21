package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SalidaDTO {
    private String fecha_hora_horario;
    private String fecha_hora_timbre;
    private Double diferencia;
    private String tipo_permiso;
    private String desde;
    private String hasta;
}
