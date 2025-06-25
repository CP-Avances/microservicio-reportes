package com.casapazmino.microservicio_reportes.model;

public class ModalidadLaboralDTO {

    private Long id;
    private String descripcion;

    public ModalidadLaboralDTO() {}

    public ModalidadLaboralDTO(Long id, String descripcion) {
        this.id = id;
        this.descripcion = descripcion;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }
}
