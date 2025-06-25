package com.casapazmino.microservicio_reportes.model;

public class CiudadDTO {

    private String provincia;
    private String nombre;

    public CiudadDTO() {}

    public CiudadDTO(String provincia, String nombre) {
        this.provincia = provincia;
        this.nombre = nombre;
    }

    public String getProvincia() {
        return provincia;
    }

    public void setProvincia(String provincia) {
        this.provincia = provincia;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
}
