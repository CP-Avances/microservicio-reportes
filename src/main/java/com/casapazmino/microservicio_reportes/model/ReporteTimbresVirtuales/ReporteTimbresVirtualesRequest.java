package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteTimbresVirtualesRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private int opcionBusqueda;
    private PeriodoDTO periodo;
    private FiltrosAplicadosDTO filtrosAplicados;
    private boolean timbreDispositivo;
    private List<GrupoTimbresDTO> data_pdf;

}
