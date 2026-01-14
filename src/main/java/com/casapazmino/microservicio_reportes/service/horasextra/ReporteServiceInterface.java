// ===== Interfaz del Servicio =====
package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;

public interface ReporteServiceInterface {
  byte[] generarReporte(ReporteHorasExtraRequest request);
}