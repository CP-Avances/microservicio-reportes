package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class MarcaAlimentacionDTO extends MarcaDTO {
    private Double minutos_alimentacion;
}
