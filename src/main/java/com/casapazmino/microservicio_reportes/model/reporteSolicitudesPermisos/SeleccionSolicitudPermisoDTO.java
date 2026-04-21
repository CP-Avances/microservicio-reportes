package com.casapazmino.microservicio_reportes.model.reporteSolicitudesPermisos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SeleccionSolicitudPermisoDTO {

    private List<Long> idsSucursales;
    private List<Long> idsRegimenes;
    private List<Long> idsDepartamentos;
    private List<Long> idsCargos;
    private List<Long> idsEmpleados;

    private List<String> nombresSucursales;
    private List<String> nombresRegimenes;
    private List<String> nombresDepartamentos;
    private List<String> nombresCargos;
    private List<String> nombresEmpleados;
}