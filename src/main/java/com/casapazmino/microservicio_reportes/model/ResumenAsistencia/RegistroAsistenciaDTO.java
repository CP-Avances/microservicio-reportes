package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RegistroAsistenciaDTO {

    private String tipo;
    private String origen;
    private Boolean control;

    private MarcaDTO entrada;
    private MarcaDTO salida;
    private MarcaAlimentacionDTO inicioAlimentacion;
    private MarcaAlimentacionDTO finAlimentacion;

    private Double minLaborados;
    private Double minPlanificados;
    private Double minAlimentacion;
    private Double minAtrasos;
    private Double minSalidasAnticipadas;

    private String observaciones;
}
