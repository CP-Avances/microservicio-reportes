package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class PlanificacionEmpleadoDTO {
    private String nombre;
    private String apellido;
    private String identificacion;
    private String codigo;
    private String departamento;
    private String cargo;
    private String ciudad;
    private String sucursal;
    private String regimen;
    private List<PlanificacionHorarioMensualDTO> horarios;
}
