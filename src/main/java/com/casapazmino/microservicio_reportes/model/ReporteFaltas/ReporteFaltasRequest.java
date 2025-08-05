package com.casapazmino.microservicio_reportes.model.ReporteFaltas;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteFaltasRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private String fechaInicio;
    private String fechaFin;
    private int opcionBusqueda;
    private boolean resumen;

    private List<GrupoFaltasDTO> grupos;

}
