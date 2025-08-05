package com.casapazmino.microservicio_reportes.model.ReporteUsuario;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class AgrupadorUsuariosDTO {

    private String sucursal;
    private String ciudad;
    private String nombre;   
    private String departamento;
    private List<UsuarioDTO> empleados;

}
