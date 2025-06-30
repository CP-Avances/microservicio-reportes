package com.casapazmino.microservicio_reportes.model.Coordenada;

public class CoordenadaDTO {
    private int id;
    private String descripcion;
    private String latitud;
    private String longitud;

    public CoordenadaDTO() {
    }

    public CoordenadaDTO(int id, String descripcion, String latitud, String longitud) {
        this.id = id;
        this.descripcion = descripcion;
        this.latitud = latitud;
        this.longitud = longitud;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }

    public String getLatitud() {
        return latitud;
    }

    public void setLatitud(String latitud) {
        this.latitud = latitud;
    }

    public String getLongitud() {
        return longitud;
    }

    public void setLongitud(String longitud) {
        this.longitud = longitud;
    }
}
