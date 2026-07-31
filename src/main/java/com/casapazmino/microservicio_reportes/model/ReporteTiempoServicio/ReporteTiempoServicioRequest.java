package com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteTiempoServicioRequest {

    private String usuario;
    private String empresa;

    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private String titulo;
    private String tipoFiltro;
    private String fechaCorte;

    private Integer opcionBusqueda;

    private ResumenTiempoServicioDTO resumen;

    private List<GrupoTiempoServicioDTO> grupos;
}