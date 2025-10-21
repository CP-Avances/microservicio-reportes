package com.casapazmino.microservicio_reportes.model.RegimenLaboral;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RangoAntiguedadDTO {
    private Integer anio_desde;
    private Integer anio_hasta;
    private Integer dias_antiguedad;
}
