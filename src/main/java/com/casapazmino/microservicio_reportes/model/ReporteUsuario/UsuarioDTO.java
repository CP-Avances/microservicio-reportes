package com.casapazmino.microservicio_reportes.model.ReporteUsuario;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class UsuarioDTO {

    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String usuario;
    private String genero;
    private String nacionalidad;
    private String ciudad;
    private String sucursal;
    private String regimen;
    private String departamento;
    private String cargo;
    private String rol;
    private String correo;
}
