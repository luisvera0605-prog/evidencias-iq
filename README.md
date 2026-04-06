# EvidenciasIQ — Flor de Tabasco

Verificación automática de evidencias de pago con IA (Claude Vision) conectado a SharePoint.

## Configuración

Las credenciales de Azure AD ya están configuradas en `src/App.tsx`:
- Client ID: `b271f29f-65f7-476e-a272-63669bdfd85e`
- Tenant ID: `746b050c-a1ff-45b9-9858-e142490982b7`
- SharePoint: `https://cisurft.sharepoint.com/sites/PlaneacionFinanciera`

## Instalación local

```bash
npm install
npm run dev
```

## Deploy en Vercel

1. Sube este proyecto a GitHub
2. Importa en vercel.com
3. Agrega la URL de Vercel en Azure AD → Authentication → URI de redirección
4. Deploy automático

## Uso

1. Iniciar sesión con cuenta Microsoft de flordetabasco.com
2. Ingresar ruta: `Ventas`
3. Clic "Escanear carpetas" — lee todos los folios VTA-XXXX
4. Clic "Analizar todo" — Claude Vision analiza cada imagen/PDF
5. Exportar CSV con resultados

## Funcionalidades

- Login OAuth2 PKCE (sin admin requerido)
- Paginación automática (sin límite de carpetas)
- Soporta imágenes (JPG, PNG, WEBP) y PDFs
- Semáforo verde/amarillo/rojo por evidencia
- Extrae: fecha, monto, referencia, cliente, banco
- Exportación CSV con todos los resultados
