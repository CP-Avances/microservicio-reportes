package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RegistroAsistenciaDTO {

    private String tipo;
    private String origen;
    private boolean control;

    private MarcaDTO entrada;
    private MarcaDTO salida;
    private MarcaAlimentacionDTO inicioAlimentacion;
    private MarcaAlimentacionDTO finAlimentacion;

    private Double minLaborados;
    private Double minPlanificados;
    private Double minAlimentacion;
    private Double minAtrasos;
    private Double minSalidasAnticipadas;
 
}