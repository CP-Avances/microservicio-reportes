package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class PlanificacionDetalleDTO {
    private String horario;
    private String entrada;
    private String inicio_comida;
    private String fin_comida;
    private String salida;
    private String entrada_;
    private String salida_;
    private String acciones;
}
