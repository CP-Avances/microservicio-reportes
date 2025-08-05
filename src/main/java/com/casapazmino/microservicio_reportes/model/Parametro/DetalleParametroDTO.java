package com.casapazmino.microservicio_reportes.model.Parametro;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DetalleParametroDTO {
    private int id;
    private String descripcion;
    private String observacion;
}
