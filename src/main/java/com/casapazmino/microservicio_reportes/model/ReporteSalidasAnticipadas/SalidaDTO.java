package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SalidaDTO {
    private String fecha_hora_horario;
    private String fecha_hora_timbre;
    private Double diferencia;
    private Double diferencia_original;
    private String tipo_permiso;
    private String desde;
    private String hasta;
    private Double diferencia_permiso;
    private Double permiso_aplicado;
    private Long id_emple_permiso;
    private String permiso_tiempo;
    private String permiso_decimal;
}