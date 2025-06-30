package com.casapazmino.microservicio_reportes.model.Discapacidad;

public class DiscapacidadDTO {

    private Long id;
    private String nombre;

    public DiscapacidadDTO() {
    }

    public DiscapacidadDTO(Long id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
}
