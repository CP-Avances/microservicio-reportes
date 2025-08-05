package com.casapazmino.microservicio_reportes.model.Regimen;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RegimenDTO {

    private Long id;
    private String descripcion;
    private String pais;
    private int mes_periodo;
    private int dias_mes;
    private int vacacion_dias_laboral;
    private int vacacion_dias_libre;
    private int vacacion_dias_calendario;
    private int dias_maximo_acumulacion;
    private int anio_antiguedad;
    private int dias_antiguedad;

    public RegimenDTO() {
    }

}
