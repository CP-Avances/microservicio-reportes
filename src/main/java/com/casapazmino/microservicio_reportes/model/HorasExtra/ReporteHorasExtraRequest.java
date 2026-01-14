package com.casapazmino.microservicio_reportes.model.HorasExtra;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteHorasExtraRequest {
    private String fechaDesde;
    private String fechaHasta;
    private String ciudad;
    private String parametro;
    private List<Long> idsSeleccionados;

    private DataResponse data;

    private String formato;
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
}

