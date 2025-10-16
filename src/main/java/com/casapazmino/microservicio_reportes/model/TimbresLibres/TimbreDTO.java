package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TimbreDTO {
    private String fechaServidor;     // yyyy-mm-dd
    private String horaServidor;      // HH:mm:ss
    private String fechaDispositivo;  // opcional según flag en FE
    private String horaDispositivo;   // opcional
    private String id_reloj;          // puede ser null
    private String accion;            // 'Timbre libre', 'Entrada', etc. (ya mapeado por FE)
    private String observacion;       // ""
    private String latitud;           // ""
    private String longitud;          // ""
}
