package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RegistroAlimentacionDTO {

    private String fecha;
    private String inicioAlimentacion;
    private String finAlimentacion;
    private String estadoInicio;
    private String estadoFin;
    private String tipoJustificacion;
    private String tipoJustificacionTexto;
    private String descripcionJustificacion;
    private String tipoPermiso;

    private String desde;
    private String hasta;

    private Boolean esJustificacionPermiso;
    private Boolean esJustificacionVacacion;
    private Boolean esJustificacionHoraExtra;
    private Double minutosPermitidos;
    private Double minutosTomados;
    private Double minutosExcesoOriginal;
    private Double minutosPermisoAplicados;
    private Double minutosHoraExtraAplicados;
    private Double minutosJustificacionAplicados;
    private Double minutosExceso;
    private Boolean calculoValido;
    private String motivoCalculo;

    private Boolean justificada;
    private Boolean tieneNovedad;
    private Boolean tieneAmbosTimbres;
    private Integer idEmplePermiso;
    private Integer idSolicitudVacacion;
}