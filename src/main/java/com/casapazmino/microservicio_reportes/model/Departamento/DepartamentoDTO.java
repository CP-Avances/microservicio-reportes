package com.casapazmino.microservicio_reportes.model.Departamento;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DepartamentoDTO {

    private Long id;
    private String nomsucursal;
    private String nombre;
    private int nivel;
    private String departamento_padre;

    public DepartamentoDTO() {}

    public DepartamentoDTO(Long id, String nomsucursal, String nombre, int nivel, String departamento_padre) {
        this.id = id;
        this.nomsucursal = nomsucursal;
        this.nombre = nombre;
        this.nivel = nivel;
        this.departamento_padre = departamento_padre;
    }
}
