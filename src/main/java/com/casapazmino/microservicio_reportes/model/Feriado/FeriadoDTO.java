package com.casapazmino.microservicio_reportes.model.Feriado;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class FeriadoDTO {
    private Integer id;
    private String descripcion;
    private String fecha;
    private String fechaRecuperacion;

}
