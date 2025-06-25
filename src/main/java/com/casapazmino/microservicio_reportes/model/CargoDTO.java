package com.casapazmino.microservicio_reportes.model;

public class CargoDTO {

    private Long id;
    private String cargo;

    public CargoDTO() {}

    public CargoDTO(Long id, String cargo) {
        this.id = id;
        this.cargo = cargo;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getCargo() {
        return cargo;
    }

    public void setCargo(String cargo) {
        this.cargo = cargo;
    }
}
