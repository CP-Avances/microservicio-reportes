package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteTimbresLibresRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String titulo;
    private String tipoFiltro;
    private Integer opcionBusqueda;
    private PeriodoDTO periodo;
    private List<DatoGrupoDTO> datos;
}
