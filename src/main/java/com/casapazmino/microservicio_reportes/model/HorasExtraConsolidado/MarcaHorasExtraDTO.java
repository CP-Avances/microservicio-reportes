package com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class MarcaHorasExtraDTO {
    private String fecha_horario;
    private String fecha_hora_horario;
    private String fecha_hora_timbre;
    private String estado_timbre;
    private Double minutos_alimentacion;
}