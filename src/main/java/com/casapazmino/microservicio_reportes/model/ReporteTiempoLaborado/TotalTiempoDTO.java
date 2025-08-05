package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TotalTiempoDTO {
    private String sucursal;
    private String nombre;
    private String formato_general_planificado;
    private String formato_decimal_planifiado;
    private String formato_general_laborado;
    private String formato_decimal_laborado;
}
