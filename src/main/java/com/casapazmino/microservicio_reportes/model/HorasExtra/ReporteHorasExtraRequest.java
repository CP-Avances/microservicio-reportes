package com.casapazmino.microservicio_reportes.model.HorasExtra;

import com.casapazmino.microservicio_reportes.model.DataGenerico;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteHorasExtraRequest extends DataGenerico {

    private List<Long> idsSeleccionados;

    private DataResponse data;


}

