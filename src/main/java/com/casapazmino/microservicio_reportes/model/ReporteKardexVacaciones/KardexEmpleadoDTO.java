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
public class KardexEmpleadoDTO {

    private Long id;
    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String ciudad;

    private Long id_suc;
    private String name_suc;
    private Long id_depa;
    private String name_dep;
    private Long id_regimen;
    private String name_regimen;
    private Long id_cargo_;
    private String name_cargo;

    // duplicados que también llegan (no estorban)
    private String sucursal;
    private String departamento;
    private String regimen;
    private String cargo;

    private String rol;
    private Boolean estado_usuario;
    private Integer min_por_dia;

    private List<KardexPeriodoDTO> periodos;
}