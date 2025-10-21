package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TotalTiempoDTO {
    private String sucursal;
    private String nombre;
    private String formato_general_planificado;
    private String formato_decimal_planifiado;
    private String formato_general_laborado;
    private String formato_decimal_laborado;
}
