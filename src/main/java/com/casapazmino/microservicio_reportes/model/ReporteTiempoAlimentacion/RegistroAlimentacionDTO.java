package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RegistroAlimentacionDTO {
    private String fecha;
    private String inicioAlimentacion;
    private String finAlimentacion;
    private Double minutosPermitidos;
    private Double minutosTomados;
    private Double minutosExceso;
}
