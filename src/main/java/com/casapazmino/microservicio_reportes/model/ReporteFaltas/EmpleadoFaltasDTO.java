package com.casapazmino.microservicio_reportes.model.ReporteFaltas;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmpleadoFaltasDTO {

    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String regimen;
    private String departamento;
    private String cargo;
    private String rol;
    private String correo;
    private int genero;
    private int id_nacionalidad;
    private List<FaltaDTO> faltas;

}