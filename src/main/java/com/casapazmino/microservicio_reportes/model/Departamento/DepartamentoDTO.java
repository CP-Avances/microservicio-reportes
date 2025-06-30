package com.casapazmino.microservicio_reportes.model.Departamento;

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

    public Long getId() {
        return id;
    }

    public String getNomsucursal() {
        return nomsucursal;
    }

    public String getNombre() {
        return nombre;
    }

    public int getNivel() {
        return nivel;
    }

    public String getDepartamento_padre() {
        return departamento_padre;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public void setNomsucursal(String nomsucursal) {
        this.nomsucursal = nomsucursal;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public void setNivel(int nivel) {
        this.nivel = nivel;
    }

    public void setDepartamento_padre(String departamento_padre) {
        this.departamento_padre = departamento_padre;
    }
}
