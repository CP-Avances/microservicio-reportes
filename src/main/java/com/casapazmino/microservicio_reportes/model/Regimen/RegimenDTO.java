package com.casapazmino.microservicio_reportes.model.Regimen;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RegimenDTO {
    private Long id;
    private String descripcion;
    private String pais;
    private Integer mes_periodo;
    private Integer dias_mes;
    private Integer vacacion_dias_laboral;
    private Integer vacacion_dias_libre;
    private Integer vacacion_dias_calendario;
    private Integer dias_maximo_acumulacion;
    private Integer anio_antiguedad;
    private Integer dias_antiguedad;
}
