package com.casapazmino.microservicio_reportes.model.Horario;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
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
}
