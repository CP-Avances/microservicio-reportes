package com.casapazmino.microservicio_reportes.model.Empleado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
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
    private String nombre;
    private String apellido;
}
