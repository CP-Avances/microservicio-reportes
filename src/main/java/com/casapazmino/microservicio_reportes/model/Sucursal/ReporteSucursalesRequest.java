package com.casapazmino.microservicio_reportes.model.Sucursal;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteSucursalesRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<SucursalDTO> sucursales;

    public ReporteSucursalesRequest() {}
    
}
