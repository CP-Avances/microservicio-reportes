package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtualesMovil;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class FiltrosAplicadosDTO {
    private Boolean bool_reg;
    private Boolean bool_dep;
    private Boolean bool_cargo;
    private Boolean bool_suc;
    private Boolean bool_emp;
}
