package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TotalAlimentacionDTO {
    private String sucursal;
    private String nombre;
    private int totalMinExceso;


}
