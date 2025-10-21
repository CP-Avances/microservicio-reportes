package com.casapazmino.microservicio_reportes.model.Parametro;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ParametroDTO {
    private Integer id;
    private String descripcion;
    private List<DetalleParametroDTO> detalles;
}
