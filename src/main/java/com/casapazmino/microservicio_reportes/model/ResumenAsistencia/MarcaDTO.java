package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

public class MarcaDTO {

    private String fecha_horario;
    private String fecha_hora_horario;
    private String fecha_hora_timbre;

    public String getFecha_horario() {
        return fecha_horario;
    }
    public void setFecha_horario(String fecha_horario) {
        this.fecha_horario = fecha_horario;
    }

    public String getFecha_hora_horario() {
        return fecha_hora_horario;
    }
    public void setFecha_hora_horario(String fecha_hora_horario) {
        this.fecha_hora_horario = fecha_hora_horario;
    }

    public String getFecha_hora_timbre() {
        return fecha_hora_timbre;
    }
    public void setFecha_hora_timbre(String fecha_hora_timbre) {
        this.fecha_hora_timbre = fecha_hora_timbre;
    }
}
