package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
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

  
    public List<VacunaUsuarioDTO> getVacunas() {
        return vacunas;
    }

    public void setVacunas(List<VacunaUsuarioDTO> vacunas) {
        this.vacunas = vacunas;
    }


}
