# CRM Facturación — Trayectoria EST

CRM generado automáticamente cada día hábil a las 8am desde SharePoint.

## Cómo funciona

1. **GitHub Actions** ejecuta `build_crm.py` cada día hábil a las 8:00 AM (Chile)
2. El script lee `Facturación EST 2025.xlsx` y `Facturación EST 2026.xlsx` desde SharePoint via Microsoft Graph API
3. Genera `index.html` con todos los datos embebidos y hace commit automático
4. **Netlify** detecta el commit y publica el sitio actualizado en minutos

## Marcar factura como pagada

Cuando se hace clic en **✓ Pagar** en el CRM:
- Se llama a `/api/pagar` (Netlify Function)
- La función actualiza la columna **G (Estado) = "Pagada"** y **I (Fecha de Pago) = fecha actual** en el Excel de SharePoint directamente via Graph API
- El cambio se refleja inmediatamente en el CRM y queda guardado en el Excel

## Secrets requeridos (GitHub + Netlify)

| Variable       | Descripción                          |
|----------------|--------------------------------------|
| `TENANT_ID`    | ID del tenant de Microsoft 365       |
| `CLIENT_ID`    | Client ID de la app Azure AD         |
| `CLIENT_SECRET`| Client Secret de la app Azure AD     |

Los mismos secrets del CRM de contratos funcionan aquí.

## Fuentes de datos

- `Administración y Finanzas/2025/Contabilidad/Trayectoria EST/Facturación/Facturación EST 2025.xlsx`
- `Administración y Finanzas/2026/Contabilidad/Trayectoria EST/Facturación/Facturación EST 2026.xlsx`
