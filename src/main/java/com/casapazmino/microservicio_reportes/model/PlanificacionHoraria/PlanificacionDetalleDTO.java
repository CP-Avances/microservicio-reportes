package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

public class PlanificacionDetalleDTO {

    private String horario;
    private String entrada;
    private String inicioComida;
    private String finComida;
    private String salida;
    private String entrada_; // formateado
    private String salida_;
    private String acciones;

    // Getters y Setters
    public String getHorario() {
        return horario;
    }

    public void setHorario(String horario) {
        this.horario = horario;
    }

    public String getEntrada() {
        return entrada;
    }

    public void setEntrada(String entrada) {
        this.entrada = entrada;
    }

    public String getInicioComida() {
        return inicioComida;
    }

    public void setInicioComida(String inicioComida) {
        this.inicioComida = inicioComida;
    }

    public String getFinComida() {
        return finComida;
    }

    public void setFinComida(String finComida) {
        this.finComida = finComida;
    }

    public String getSalida() {
        return salida;
    }

    public void setSalida(String salida) {
        this.salida = salida;
    }

    public String getEntrada_() {
        return entrada_;
    }

    public void setEntrada_(String entrada_) {
        this.entrada_ = entrada_;
    }

    public String getSalida_() {
        return salida_;
    }

    public void setSalida_(String salida_) {
        this.salida_ = salida_;
    }

    public String getAcciones() {
        return acciones;
    }

    public void setAcciones(String acciones) {
        this.acciones = acciones;
    }
}
