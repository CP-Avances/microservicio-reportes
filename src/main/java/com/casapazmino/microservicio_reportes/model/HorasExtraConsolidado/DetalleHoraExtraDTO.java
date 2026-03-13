package com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DetalleHoraExtraDTO {
    private Double minutosHorasExtra;
    private String horasExtra;
    private String tipoPorcentaje;
    private String porcentaje;
    private String tipoRecargo;
}