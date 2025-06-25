package com.casapazmino.microservicio_reportes.model;

public class VacunaUsuarioDTO {

    private String tipo_vacuna;
    private String fecha;
    private String descripcion;

    public String getTipo_vacuna() {
        return tipo_vacuna;
    }

    public void setTipo_vacuna(String tipo_vacuna) {
        this.tipo_vacuna = tipo_vacuna;
    }

    public String getFecha() {
        return fecha;
    }

    public void setFecha(String fecha) {
        this.fecha = fecha;
    }

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }
}
