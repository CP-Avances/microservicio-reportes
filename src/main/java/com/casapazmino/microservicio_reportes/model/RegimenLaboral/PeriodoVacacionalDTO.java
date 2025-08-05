package com.casapazmino.microservicio_reportes.model.RegimenLaboral;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class PeriodoVacacionalDTO {
    private String descripcion;
    private Integer dias_vacacion;
}