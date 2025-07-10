package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RegistroAlimentacionDTO {
    private String fecha;
    private String inicioAlimentacion;
    private String finAlimentacion;
    private double minutosPermitidos;
    private double minutosTomados;
    private double minutosExceso;

}
