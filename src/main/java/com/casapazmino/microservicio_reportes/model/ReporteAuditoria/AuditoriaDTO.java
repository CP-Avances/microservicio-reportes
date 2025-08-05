package com.casapazmino.microservicio_reportes.model.ReporteAuditoria;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class AuditoriaDTO {
    private String plataforma;
    private String user_name;
    private String ip_address;
    private String table_name;
    private String action;
    private String fecha_hora_format;
    private String solo_hora;
    private String original_data;
    private String new_data;

}
