package com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ContratoTiempoServicioDTO {

    // ===================== IDENTIFICADORES =====================

    @JsonProperty("id_contrato")
    private Long idContrato;

    @JsonProperty("id_regimen")
    private Long idRegimen;

    @JsonProperty("id_cargo")
    private Long idCargo;

    @JsonProperty("id_tipo_cargo")
    private Long idTipoCargo;

    @JsonProperty("id_departamento")
    private Long idDepartamento;

    @JsonProperty("id_sucursal")
    private Long idSucursal;

    // ===================== NOMBRES PARA PDF Y EXCEL =====================

    private String regimen;
    private String cargo;
    private String departamento;
    private String sucursal;

    // ===================== FECHAS DEL CONTRATO =====================

    @JsonProperty("fecha_ingreso")
    private String fechaIngreso;

    @JsonProperty("fecha_salida")
    private String fechaSalida;

    @JsonProperty("fecha_fin_calculo")
    private String fechaFinCalculo;

    // ===================== RESULTADO DEL CÁLCULO =====================

    @JsonProperty("anios_servicio")
    private Integer aniosServicio;

    @JsonProperty("meses_servicio")
    private Integer mesesServicio;

    @JsonProperty("dias_servicio")
    private Integer diasServicio;

    @JsonProperty("tiempo_servicio")
    private String tiempoServicio;
}