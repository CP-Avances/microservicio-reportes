package com.casapazmino.microservicio_reportes.model;

public class NacionalidadDTO {

    private int id;
    private String nombre;

    public NacionalidadDTO() {
    }

    public NacionalidadDTO(int id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
}
