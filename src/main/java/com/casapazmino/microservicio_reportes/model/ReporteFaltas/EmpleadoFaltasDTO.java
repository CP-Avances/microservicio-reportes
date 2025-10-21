package com.casapazmino.microservicio_reportes.model.ReporteFaltas;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
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
    private Integer genero;
    private Integer id_nacionalidad;
    private String ciudad;
    private String sucursal;
    private String generoNombre;
    private String nacionalidadNombre;
    private List<FaltaDTO> faltas;
}