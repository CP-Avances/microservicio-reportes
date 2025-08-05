package com.casapazmino.microservicio_reportes.model.Horario;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class HorarioDTO {

    private String nombre;
    private String horaTrabajo;
    private int minutosComida;
    private String codigo;
    private boolean noturno;
    private String documento;
    private List<DetalleHorarioDTO> detalles;

    public HorarioDTO() {}

}
