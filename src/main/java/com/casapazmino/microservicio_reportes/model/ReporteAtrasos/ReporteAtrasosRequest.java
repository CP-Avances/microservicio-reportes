package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class ReporteAtrasosRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private String fechaInicio;
    private String fechaFin;
    private String opcionBusqueda;

    private FiltroResumenDTO resumen;
    private List<TotalAtrasosDTO> totales;
    private List<GrupoAtrasoDTO> grupos;
}