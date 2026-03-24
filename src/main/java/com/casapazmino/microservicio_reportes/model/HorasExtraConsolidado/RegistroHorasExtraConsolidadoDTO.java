package com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RegistroHorasExtraConsolidadoDTO {

    private String tipo;
    private String origen;
    private Boolean control;

    private MarcaHorasExtraDTO entrada;
    private MarcaHorasExtraDTO salida;
    private MarcaAlimentacionHorasExtraDTO inicioAlimentacion;
    private MarcaAlimentacionHorasExtraDTO finAlimentacion;

    private Double minLaborados;
    private Double minPlanificados;
    private Double minAlimentacion;
    private Double minAtrasos;
    private Double minSalidasAnticipadas;

    private String horasExtra;
    private Double minutosHorasExtra;
    private String tipoPorcentaje;
    private String porcentaje;
    private String tipoRecargo;

    private List<DetalleHoraExtraDTO> detalleHorasExtra;
}