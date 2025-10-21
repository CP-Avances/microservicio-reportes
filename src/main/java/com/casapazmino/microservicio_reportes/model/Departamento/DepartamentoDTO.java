package com.casapazmino.microservicio_reportes.model.Departamento;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DepartamentoDTO {
    private Long id;
    private String nomsucursal;
    private Integer id_sucursal;
    private String nombre;
    private Integer nivel;
    private String departamento_padre;
}

