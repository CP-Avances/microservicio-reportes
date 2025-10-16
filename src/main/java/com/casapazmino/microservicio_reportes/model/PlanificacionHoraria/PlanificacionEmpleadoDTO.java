package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
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
