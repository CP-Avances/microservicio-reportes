package com.casapazmino.microservicio_reportes.model.HorasExtra;

import lombok.Data;

import java.util.List;

@Data
public class DataResponse {
    private Boolean ok;
    private List<DataHoraExtraLista> data;
}
