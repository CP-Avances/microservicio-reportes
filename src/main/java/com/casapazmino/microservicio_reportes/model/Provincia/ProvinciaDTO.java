package com.casapazmino.microservicio_reportes.model.Provincia;

public class ProvinciaDTO {

    private String pais;
    private String nombre;

    public ProvinciaDTO() {}

    public ProvinciaDTO(String pais, String nombre) {
        this.pais = pais;
        this.nombre = nombre;
    }

    public String getPais() {
        return pais;
    }

    public void setPais(String pais) {
        this.pais = pais;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
}
