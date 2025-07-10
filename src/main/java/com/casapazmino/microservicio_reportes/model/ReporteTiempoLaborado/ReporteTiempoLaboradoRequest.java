package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import lombok.Getter;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
public class ReporteTiempoLaboradoRequest {
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

    private List<TotalTiempoDTO> totales;
    private List<GrupoTiempoDTO> grupos;

    public static class ResumenResumen {
        private boolean bool_reg;
        private boolean bool_dep;
        private boolean bool_cargo;
        private boolean bool_suc;
        private boolean bool_emp;

    }
}
