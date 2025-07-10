package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RegistroTiempoDTO {
    private String tipo;
    private String origen;
    private Boolean control;

    private CampoTiempoDTO entrada;
    private CampoTiempoDTO salida;
    private CampoTiempoDTO inicioAlimentacion;
    private CampoTiempoDTO finAlimentacion;

    private String minPlanificados;
    private String minLaborados;
    private String tiempoPlanificado;
    private String tiempoLaborado;
}
