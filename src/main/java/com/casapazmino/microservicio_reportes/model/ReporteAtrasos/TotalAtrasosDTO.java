package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TotalAtrasosDTO {
    private String sucursal;
    private String nombre;
    private String formato_general;
    private String formato_decimal;
}
