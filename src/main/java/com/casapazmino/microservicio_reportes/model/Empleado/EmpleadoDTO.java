package com.casapazmino.microservicio_reportes.model.Empleado;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmpleadoDTO {

    private String codigo;
    private String nombreCompleto;
    private String identificacion;
    private String fechaNacimiento;
    private String correo;
    private String genero;
    private String estadoCivil;
    private String domicilio;
    private String telefono;
    private String estadoTexto;
    private String nacionalidad;

    public EmpleadoDTO() {}

}
