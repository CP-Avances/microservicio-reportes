package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class VacunaUsuarioDTO {

    private String tipo_vacuna;
    private String fecha;
    private String descripcion;
    private String carnet;

}
