package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TimbreDTO {
    private String fechaHora;
    private String fechaHorario;
    private String horaHorario;
    private String horaTimbre;
    private String accion;
    private String accionTexto;
    private String estadoTimbre;
    private String tipoDia;
    private String estadoOrigen;
    private Boolean justificada;
    private String estadoJustificacion;
    private Boolean tienePermiso;
    private Boolean tieneVacacion;
    private String tipoJustificacion;
    private String nombreJustificacion;
    private String desde;
    private String hasta;
    private String detalleJustificacion;
    private String tipoPermiso;
    private String permisoDesde;
    private String permisoHasta;
    private String idEmplePermiso;
    private String vacacionDesde;
    private String vacacionHasta;
    private String numeroDiasVacacion;
    private String idSolicitudVacacion;
}