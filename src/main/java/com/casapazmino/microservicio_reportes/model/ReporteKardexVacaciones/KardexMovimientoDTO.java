package com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class KardexMovimientoDTO {

    private String tipo; // "PERMISO" | "VACACION"
    private Long id_mov;
    private String detalle;

    private String fecha_inicio;
    private String fecha_final;

    private String hora_inicio; // puede ser null
    private String hora_fin;    // puede ser null

    private Double descuento_decimal;
    private DHMDTO descuento_dhm;

    private SemanaDHMDTO semana_descuento;

    private Double saldo_decimal;
    private DHMDTO saldo_dhm;
}