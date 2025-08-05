package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import java.util.List;
import java.util.Map;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteAsistenciaRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String fechaInicio;
    private String fechaFin;
    private int opcionBusqueda;
    private Map<String, Boolean> resumen;
    private List<GrupoAsistenciaDTO> grupos;
 
}