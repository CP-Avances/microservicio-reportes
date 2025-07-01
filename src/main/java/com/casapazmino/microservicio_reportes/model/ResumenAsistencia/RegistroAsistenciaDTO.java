package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

public class RegistroAsistenciaDTO {

    private String tipo;
    private String origen;
    private boolean control;

    private MarcaDTO entrada;
    private MarcaDTO salida;
    private MarcaAlimentacionDTO inicioAlimentacion;
    private MarcaAlimentacionDTO finAlimentacion;

    private int minLaborados;
    private int minPlanificados;
    private int minAlimentacion;
    private int minAtrasos;
    private int minSalidasAnticipadas;

    public String getTipo() {
        return tipo;
    }
    public void setTipo(String tipo) {
        this.tipo = tipo;
    }

    public String getOrigen() {
        return origen;
    }
    public void setOrigen(String origen) {
        this.origen = origen;
    }

    public boolean isControl() {
        return control;
    }
    public void setControl(boolean control) {
        this.control = control;
    }

    public MarcaDTO getEntrada() {
        return entrada;
    }
    public void setEntrada(MarcaDTO entrada) {
        this.entrada = entrada;
    }

    public MarcaDTO getSalida() {
        return salida;
    }
    public void setSalida(MarcaDTO salida) {
        this.salida = salida;
    }

    public MarcaAlimentacionDTO getInicioAlimentacion() {
        return inicioAlimentacion;
    }
    public void setInicioAlimentacion(MarcaAlimentacionDTO inicioAlimentacion) {
        this.inicioAlimentacion = inicioAlimentacion;
    }

    public MarcaAlimentacionDTO getFinAlimentacion() {
        return finAlimentacion;
    }
    public void setFinAlimentacion(MarcaAlimentacionDTO finAlimentacion) {
        this.finAlimentacion = finAlimentacion;
    }

    public int getMinLaborados() {
        return minLaborados;
    }
    public void setMinLaborados(int minLaborados) {
        this.minLaborados = minLaborados;
    }

    public int getMinPlanificados() {
        return minPlanificados;
    }
    public void setMinPlanificados(int minPlanificados) {
        this.minPlanificados = minPlanificados;
    }

    public int getMinAlimentacion() {
        return minAlimentacion;
    }
    public void setMinAlimentacion(int minAlimentacion) {
        this.minAlimentacion = minAlimentacion;
    }

    public int getMinAtrasos() {
        return minAtrasos;
    }
    public void setMinAtrasos(int minAtrasos) {
        this.minAtrasos = minAtrasos;
    }

    public int getMinSalidasAnticipadas() {
        return minSalidasAnticipadas;
    }
    public void setMinSalidasAnticipadas(int minSalidasAnticipadas) {
        this.minSalidasAnticipadas = minSalidasAnticipadas;
    }
}
