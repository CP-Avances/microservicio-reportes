package com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class KardexPendienteDTO {

    private String tipo; // "PERMISO" | "VACACION"
    private Long id_mov;

    private Long id_empleado;
    private Long id_periodo_vacacion;

    private String detalle;

    private String fecha_inicio;
    private String fecha_final;

    private String hora_inicio; // puede ser null
    private String hora_fin;    // puede ser null

    // OJO: en tu JSON viene como "0" (string), por eso lo dejamos String
    private String dias;

    private Integer minutos_totales;

    private Integer estado;
    private String estado_texto;
}