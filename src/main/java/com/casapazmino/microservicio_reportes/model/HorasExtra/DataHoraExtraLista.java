package com.casapazmino.microservicio_reportes.model.HorasExtra;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.LocalTime;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class DataHoraExtraLista {
    private Long idEmpleado;
    private String identificacion;
    private String codigoEmpleado;
    private String nombreCompleto;
    private Long idCalculoHoraExtra;

    @JsonFormat(pattern = "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'")
    private LocalDateTime fecha;

    @JsonFormat(pattern = "HH:mm:ss")
    private LocalTime horaEntradaReal;

    @JsonFormat(pattern = "HH:mm:ss")
    private LocalTime horaSalidaReal;

    private Integer minutosHorasExtra;
    private BigDecimal horasHorasExtra;
    private BigDecimal totalAPagar;
    private String estadoCalculoDesc;
}