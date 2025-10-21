package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TotalAlimentacionDTO {
    private String sucursal;
    private String nombre;
    private Integer totalMinExceso;
}
