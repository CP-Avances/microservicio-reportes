package com.casapazmino.microservicio_reportes.model.ResumenGeneral;

import com.casapazmino.microservicio_reportes.model.HorasExtra.DataResponse;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Data
@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ResumenGeneral {
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
