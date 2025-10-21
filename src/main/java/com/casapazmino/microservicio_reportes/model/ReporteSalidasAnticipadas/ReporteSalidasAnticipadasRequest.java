package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteSalidasAnticipadasRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String fechaInicio;
    private String fechaFin;
    private Integer opcionBusqueda;
    private ResumenResumen resumen;
    private List<GrupoSalidasDTO> grupos;

    @Getter
    @Setter
    @NoArgsConstructor
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class ResumenResumen {
        private Boolean bool_reg;
        private Boolean bool_dep;
        private Boolean bool_cargo;
        private Boolean bool_suc;
        private Boolean bool_emp;
    }
}
