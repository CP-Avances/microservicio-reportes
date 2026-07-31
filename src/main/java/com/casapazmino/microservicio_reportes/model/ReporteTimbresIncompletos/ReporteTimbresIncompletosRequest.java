package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteTimbresIncompletosRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private Integer opcionBusqueda;
    private String filtroJustificacion;
    private PeriodoDTO periodo;
    private FiltrosAplicadosDTO filtrosAplicados;
    private List<TimbresSucursalDTO> data_pdf;
}