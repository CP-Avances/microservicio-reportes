package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteTiempoAlimentacionRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private String fechaInicio;
    private String fechaFin;
    private String opcionBusqueda;
    private ResumenResumen resumen;

    private List<TotalAlimentacionDTO> totales;
    private List<GrupoAlimentacionDTO> grupos;

    
    public static class ResumenResumen {
        private boolean bool_reg;
        private boolean bool_dep;
        private boolean bool_cargo;
        private boolean bool_suc;
        private boolean bool_emp;

    }


}
