package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class PeriodoDTO {
    private String inicio;
    private String fin;
}
