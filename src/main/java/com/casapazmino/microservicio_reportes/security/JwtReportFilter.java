package com.casapazmino.microservicio_reportes.security;

import io.jsonwebtoken.Claims;
import io.jsonwebtoken.JwtException;
import io.jsonwebtoken.Jwts;
import io.jsonwebtoken.security.Keys;
import jakarta.servlet.Filter;
import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.ServletRequest;
import jakarta.servlet.ServletResponse;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.Key;

@Component
public class JwtReportFilter implements Filter {

    @Value("${jwt.secret}")
    private String jwtSecret;

    @Override
    public void doFilter(ServletRequest request,
                         ServletResponse response,
                         FilterChain chain) throws IOException, ServletException {

        HttpServletRequest httpReq = (HttpServletRequest) request;
        HttpServletResponse httpRes = (HttpServletResponse) response;

        String method = httpReq.getMethod();

        // 1) Dejar pasar SIEMPRE la preflight OPTIONS (no se valida token aquí)
        if ("OPTIONS".equalsIgnoreCase(method)) {
            chain.doFilter(request, response);
            return;
        }

        // 2) Para todos los demás métodos (POST, GET, etc.), validar JWT
        String authHeader = httpReq.getHeader("Authorization");

        if (authHeader == null || !authHeader.startsWith("Bearer ")) {
            httpRes.setStatus(HttpServletResponse.SC_UNAUTHORIZED);
            return;
        }

        String token = authHeader.substring(7); // quitar "Bearer "

        try {
            Key key = Keys.hmacShaKeyFor(jwtSecret.getBytes(StandardCharsets.UTF_8));

            Claims claims = Jwts.parserBuilder()
                    .setSigningKey(key)
                    .build()
                    .parseClaimsJws(token)   // valida firma + expiración
                    .getBody();

            Boolean webAccess = claims.get("_web_access", Boolean.class);
            if (webAccess == null || !webAccess) {
                httpRes.setStatus(HttpServletResponse.SC_FORBIDDEN);
                return;
            }

            // Opcional: dejar datos del usuario por si quieres usarlos
            httpReq.setAttribute("userId", claims.get("_id"));
            httpReq.setAttribute("userIdEmpleado", claims.get("_id_empleado"));
            httpReq.setAttribute("idEmpresa", claims.get("_empresa"));
            httpReq.setAttribute("usuario", claims.get("_usuario"));
            httpReq.setAttribute("modulos", claims.get("modulos"));

            // Todo OK → sigue la cadena hasta los controladores
            chain.doFilter(request, response);

        } catch (JwtException ex) {
            System.out.println("==== ERROR JWT EN MICRO REPORTE ====");
            System.out.println("Mensaje: " + ex.getMessage());
            ex.printStackTrace(); // SOLO para desarrollo
            httpRes.setStatus(HttpServletResponse.SC_UNAUTHORIZED);
        }
    }
}
