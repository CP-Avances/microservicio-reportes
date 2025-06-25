package com.casapazmino.microservicio_reportes.model;

public class FeriadoDTO {
    private Long id;
    private String descripcion;
    private String fecha;
    private String fechaRecuperacion;

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

    public String getFecha() {
        return fecha;
    }

    public void setFecha(String fecha) {
        this.fecha = fecha;
    }

    public String getFechaRecuperacion() {
        return fechaRecuperacion;
    }

    public void setFechaRecuperacion(String fechaRecuperacion) {
        this.fechaRecuperacion = fechaRecuperacion;
    }
}
