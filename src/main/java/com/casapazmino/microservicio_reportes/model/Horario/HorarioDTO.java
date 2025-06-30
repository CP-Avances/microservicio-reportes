package com.casapazmino.microservicio_reportes.model.Horario;

import java.util.List;

public class HorarioDTO {

    private String nombre;
    private String horaTrabajo;
    private int minutosComida;
    private String codigo;
    private boolean noturno;
    private String documento;
    private List<DetalleHorarioDTO> detalles;

    public HorarioDTO() {}

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getHoraTrabajo() {
        return horaTrabajo;
    }

    public void setHoraTrabajo(String horaTrabajo) {
        this.horaTrabajo = horaTrabajo;
    }

    public int getMinutosComida() {
        return minutosComida;
    }

    public void setMinutosComida(int minutosComida) {
        this.minutosComida = minutosComida;
    }

    public String getCodigo() {
        return codigo;
    }

    public void setCodigo(String codigo) {
        this.codigo = codigo;
    }

    public boolean isNoturno() {
        return noturno;
    }

    public void setNoturno(boolean noturno) {
        this.noturno = noturno;
    }

    public String getDocumento() {
        return documento;
    }

    public void setDocumento(String documento) {
        this.documento = documento;
    }

    public List<DetalleHorarioDTO> getDetalles() {
        return detalles;
    }

    public void setDetalles(List<DetalleHorarioDTO> detalles) {
        this.detalles = detalles;
    }
}
