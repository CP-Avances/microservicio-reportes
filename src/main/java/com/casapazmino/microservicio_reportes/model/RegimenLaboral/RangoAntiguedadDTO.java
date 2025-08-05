package com.casapazmino.microservicio_reportes.model.RegimenLaboral;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RangoAntiguedadDTO {
    private Integer anio_desde;
    private Integer anio_hasta;
    private Integer dias_antiguedad;
}
