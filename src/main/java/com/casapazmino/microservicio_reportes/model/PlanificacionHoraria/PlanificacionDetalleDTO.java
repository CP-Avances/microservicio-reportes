package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
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
