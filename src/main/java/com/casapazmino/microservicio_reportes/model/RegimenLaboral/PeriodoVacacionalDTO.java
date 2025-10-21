package com.casapazmino.microservicio_reportes.model.RegimenLaboral;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class PeriodoVacacionalDTO {
    private String descripcion;
    private Integer dias_vacacion;
}
