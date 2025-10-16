package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import java.util.List;
import lombok.Getter;
import lombok.Setter;


@Getter
@Setter
public class EmpleadoDTO {
    private String identificacion;
    private String nombre;
    private String apellido;
    private String correo;          // puede ser null
    private String genero;          // 'Masculino', 'Femenino', 'No especificado', etc.
    private String nacionalidad;    // nombre nacionalidad
    private String cargo;           // puede ser null
    private String regimen;         // puede ser null
    private String codigo;          // puede ser null
    private String rol;             // puede ser null
    private String departamento;    // override por empleado si viene
    private String ciudad;          // override por empleado si viene
    private String sucursal;        // override por empleado si viene

    private List<TimbreDTO> timbres;
}
