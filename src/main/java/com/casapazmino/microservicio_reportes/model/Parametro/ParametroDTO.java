package com.casapazmino.microservicio_reportes.model.Parametro;

import java.util.List;

public class ParametroDTO {
    private Long id;
    private String descripcion;
    private List<ParametroDetalleDTO> detalles;

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

    public List<ParametroDetalleDTO> getDetalles() {
        return detalles;
    }

    public void setDetalles(List<ParametroDetalleDTO> detalles) {
        this.detalles = detalles;
    }
}
