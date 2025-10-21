package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class EmpleadoVacunaUsuarioDTO {

    private String identificacion;
    private String nombre;
    private String apellido;
    private String correo;
    private String genero;
    private String nacionalidad;
    private String cargo;
    private String regimen;
    private String codigo;
    private String rol;
    private String departamento;
    private String titulo;
    private String ciudad;
    private String sucursal;
    private List<VacunaUsuarioDTO> vacunas;
}
