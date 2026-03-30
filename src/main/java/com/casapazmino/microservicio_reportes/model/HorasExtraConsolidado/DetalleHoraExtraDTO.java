package com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DetalleHoraExtraDTO {

    @JsonProperty("minutosHorasExtra")
    private Double minutosHorasExtra;

    @JsonProperty("horasExtra")
    private String horasExtra;

    @JsonProperty("tipoPorcentaje")
    private String tipoPorcentaje;

    @JsonProperty("porcentaje")
    private String porcentaje;

    @JsonProperty("tipoRecargo")
    private String tipoRecargo;

    @JsonProperty("totalAPagar")
    private Double totalAPagar;

    @JsonProperty("estadoEntradaReporte")
    private String estadoEntradaReporte;

    @JsonProperty("estadoFinAlimentacionReporte")
    private String estadoFinAlimentacionReporte;

    @JsonProperty("esSoloReporte")
    private Boolean esSoloReporte;
}