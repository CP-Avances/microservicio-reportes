package com.casapazmino.microservicio_reportes.model.ReporteTimbresUsuarios;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class FiltrosAplicadosDTO {
    private boolean bool_reg;
    private boolean bool_dep;
    private boolean bool_cargo;
    private boolean bool_suc;
    private boolean bool_emp;

}
