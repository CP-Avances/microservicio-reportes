package com.casapazmino.microservicio_reportes.model.ResumenGeneral;

import com.casapazmino.microservicio_reportes.model.DataGenerico;
import com.casapazmino.microservicio_reportes.model.HorasExtra.DataResponse;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Data
@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ResumenGeneral extends DataGenerico {


    private List<Long> idsSeleccionados;

    private DataResponse data;


}
