package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class VacunaUsuarioDTO {

    private String tipo_vacuna;
    private String fecha;
    private String descripcion;
    private String carnet;
}
