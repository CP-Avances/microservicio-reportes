package com.casapazmino.microservicio_reportes.model;

public class RegimenDTO {

    private Long id;
    private String descripcion;
    private String pais;
    private int mes_periodo;
    private int dias_mes;
    private int vacacion_dias_laboral;
    private int vacacion_dias_libre;
    private int vacacion_dias_calendario;
    private int dias_maximo_acumulacion;
    private int anio_antiguedad;
    private int dias_antiguedad;

    public RegimenDTO() {
    }

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

    public String getPais() {
        return pais;
    }

    public void setPais(String pais) {
        this.pais = pais;
    }

    public int getMes_periodo() {
        return mes_periodo;
    }

    public void setMes_periodo(int mes_periodo) {
        this.mes_periodo = mes_periodo;
    }

    public int getDias_mes() {
        return dias_mes;
    }

    public void setDias_mes(int dias_mes) {
        this.dias_mes = dias_mes;
    }

    public int getVacacion_dias_laboral() {
        return vacacion_dias_laboral;
    }

    public void setVacacion_dias_laboral(int vacacion_dias_laboral) {
        this.vacacion_dias_laboral = vacacion_dias_laboral;
    }

    public int getVacacion_dias_libre() {
        return vacacion_dias_libre;
    }

    public void setVacacion_dias_libre(int vacacion_dias_libre) {
        this.vacacion_dias_libre = vacacion_dias_libre;
    }

    public int getVacacion_dias_calendario() {
        return vacacion_dias_calendario;
    }

    public void setVacacion_dias_calendario(int vacacion_dias_calendario) {
        this.vacacion_dias_calendario = vacacion_dias_calendario;
    }

    public int getDias_maximo_acumulacion() {
        return dias_maximo_acumulacion;
    }

    public void setDias_maximo_acumulacion(int dias_maximo_acumulacion) {
        this.dias_maximo_acumulacion = dias_maximo_acumulacion;
    }

    public int getAnio_antiguedad() {
        return anio_antiguedad;
    }

    public void setAnio_antiguedad(int anio_antiguedad) {
        this.anio_antiguedad = anio_antiguedad;
    }

    public int getDias_antiguedad() {
        return dias_antiguedad;
    }

    public void setDias_antiguedad(int dias_antiguedad) {
        this.dias_antiguedad = dias_antiguedad;
    }
}
