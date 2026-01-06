package com.casapazmino.microservicio_reportes.model.RegimenLaboral;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RegimenDTO {
    private String id; 
    private String descripcion;
    private String pais;
    private Boolean continuidad_laboral;
    private Integer mes_periodo;
    private Integer dias_mes;
    private Integer trabajo_minimo_mes;
    private Integer trabajo_minimo_horas;
    private Boolean antiguedad;

    private Integer vacacion_dias_laboral;
    private Integer vacacion_dias_libre;
    private Integer vacacion_dias_calendario;
    private Boolean acumular;
    private Integer dias_maximo_acumulacion;
    private Boolean vacacion_divisible;

    private Double vacacion_dias_laboral_mes;
    private Double vacacion_dias_calendario_mes;
    private Double laboral_dias;
    private Double calendario_dias;

    private Boolean antiguedad_fija;
    private Integer anio_antiguedad;
    private Integer dias_antiguedad;

    private Boolean antiguedad_variable;
    private List<RangoAntiguedadDTO> rangos_antiguedad;
}
