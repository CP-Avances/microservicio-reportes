package com.casapazmino.microservicio_reportes.model.Horario;

public class DetalleHorarioDTO {

    private int orden;
    private String hora;
    private Integer tolerancia;
    private String tipoAccionShow;
    private boolean segundoDia;
    private int minutosAntes;
    private int minutosDespues;

    public DetalleHorarioDTO() {
    }

    public int getOrden() {
        return orden;
    }

    public void setOrden(int orden) {
        this.orden = orden;
    }

    public String getHora() {
        return hora;
    }

    public void setHora(String hora) {
        this.hora = hora;
    }

    public Integer getTolerancia() {
        return tolerancia;
    }

    public void setTolerancia(Integer tolerancia) {
        this.tolerancia = tolerancia;
    }

    public String getTipoAccionShow() {
        return tipoAccionShow;
    }

    public void setTipoAccionShow(String tipoAccionShow) {
        this.tipoAccionShow = tipoAccionShow;
    }

    public boolean isSegundoDia() {
        return segundoDia;
    }

    public void setSegundoDia(boolean segundoDia) {
        this.segundoDia = segundoDia;
    }

    public int getMinutosAntes() {
        return minutosAntes;
    }

    public void setMinutosAntes(int minutosAntes) {
        this.minutosAntes = minutosAntes;
    }

    public int getMinutosDespues() {
        return minutosDespues;
    }

    public void setMinutosDespues(int minutosDespues) {
        this.minutosDespues = minutosDespues;
    }
}
