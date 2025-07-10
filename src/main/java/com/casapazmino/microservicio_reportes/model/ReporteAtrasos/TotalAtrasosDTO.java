package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TotalAtrasosDTO {
    private String sucursal;
    private String nombre;
    private String formato_general;
    private String formato_decimal;
}
