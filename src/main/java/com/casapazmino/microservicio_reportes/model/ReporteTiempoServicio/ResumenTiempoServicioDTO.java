package com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ResumenTiempoServicioDTO {

    private Boolean bool_reg;
    private Boolean bool_dep;
    private Boolean bool_cargo;
    private Boolean bool_suc;
    private Boolean bool_emp;

    private Boolean bool_tab;
    private Boolean bool_inc;
}