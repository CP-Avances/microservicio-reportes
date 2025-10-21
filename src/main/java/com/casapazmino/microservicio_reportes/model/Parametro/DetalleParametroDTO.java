package com.casapazmino.microservicio_reportes.model.Parametro;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DetalleParametroDTO {
    private Integer id;
    private String descripcion;
    private String observacion;
}
