package com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class KardexPeriodoDTO {

    private Long id_periodo_vacacion;
    private String estado; // "ACTIVO" | "INACTIVO"

    // fechas
    private String fecha_inicio;
    private String fecha_final;
    private String fecha_desde;
    private String fecha_acreditar_vacaciones;
    private String fecha_ultima_actualizacion;

    // fila PDF
    private String f_ingreso;
    private String f_salida;
    private String f_carga;

    private Double dias_acum_decimal;
    private DHMDTO dias_acum_dhm;

    private String estado_periodo; // "ACTIVO" | "INACTIVO"
    private Integer dias_antiguedad;

    // decimales + dhm
    private Double proporcional_decimal;
    private DHMDTO proporcional_dhm;

    private Double liquidacion_decimal;
    private DHMDTO liquidacion_dhm;

    private Double saldo_transferido_decimal;
    private DHMDTO saldo_transferido_dhm;

    private Double saldo_inicial_decimal;
    private DHMDTO saldo_inicial_dhm;

    private Double saldo_disponible_decimal;
    private DHMDTO saldo_disponible_dhm;

    // otros campos que llegan del backend (por si se usan en el reporte)
    private Double dias_vacacion;
    private Double usados_dias_vacacion;
    private Double usados_antiguedad;
    private String observacion;

    private List<KardexMovimientoDTO> movimientos;
    private List<KardexPendienteDTO> pendientes;
}