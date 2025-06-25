package com.casapazmino.microservicio_reportes.model;

public class TituloDTO {

    private int id;
    private String nivel;
    private String nombre;

    public TituloDTO() {
    }

    public TituloDTO(int id, String nivel, String nombre) {
        this.id = id;
        this.nivel = nivel;
        this.nombre = nombre;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNivel() {
        return nivel;
    }

    public void setNivel(String nivel) {
        this.nivel = nivel;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
}
