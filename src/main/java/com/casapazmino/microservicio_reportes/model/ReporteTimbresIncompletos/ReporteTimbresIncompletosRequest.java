package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class ReporteTimbresIncompletosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private int opcionBusqueda;

    private PeriodoDTO periodo;
    private FiltrosAplicadosDTO filtrosAplicados;
    private List<TimbresSucursalDTO> data_pdf;
    

}
