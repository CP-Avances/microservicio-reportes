package com.casapazmino.microservicio_reportes.model.Parametro;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ParametroDTO {
    private int id;
    private String descripcion;
    private List<DetalleParametroDTO> detalles;
}
