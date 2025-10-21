package com.casapazmino.microservicio_reportes.model.Feriado;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class FeriadoDTO {
    private Integer id;
    private String descripcion;
    private String fecha;
    private String fechaRecuperacion;
}
