package com.casapazmino.microservicio_reportes.model.EstadoCivil;

public class EstadoCivilDTO {

    private int id;
    private String estadoCivil;

    public EstadoCivilDTO() {
    }

    public EstadoCivilDTO(int id, String estadoCivil) {
        this.id = id;
        this.estadoCivil = estadoCivil;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getEstadoCivil() {
        return estadoCivil;
    }

    public void setEstadoCivil(String estadoCivil) {
        this.estadoCivil = estadoCivil;
    }
}
