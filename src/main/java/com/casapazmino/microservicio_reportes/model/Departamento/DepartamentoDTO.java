package com.casapazmino.microservicio_reportes.model.Departamento;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DepartamentoDTO {

    private Long id;
    private String nomsucursal;
    private Integer id_sucursal;
    private String nombre;
    private Integer nivel;
    private String departamento_padre;

    public DepartamentoDTO() {}

    public DepartamentoDTO(Long id, String nomsucursal, String nombre, Integer nivel, String departamento_padre) {
        this.id = id;
        this.nomsucursal = nomsucursal;
        this.nombre = nombre;
        this.nivel = nivel;
        this.departamento_padre = departamento_padre;
    }
}
