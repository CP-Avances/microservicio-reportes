package com.casapazmino.microservicio_reportes.model.Feriado;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class FeriadoDTO {
    private Long id;
    private String descripcion;
    private String fecha;
    private String fechaRecuperacion;

}
