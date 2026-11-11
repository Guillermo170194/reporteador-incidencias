# Reporteador IMSS-BIENESTAR — versión modular inicial

## Ejecución local

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

## Render
- Usar servicio web Python y variables de entorno: `SUPABASE_URL`, `SUPABASE_KEY`, `GOOGLE_CREDENTIALS` (JSON completo de cuenta de servicio).
- Comando de inicio: `streamlit run app.py --server.address 0.0.0.0 --server.port $PORT --server.headless true`
- No incluye usuario ni contraseña en la aplicación. **Importante:** la URL pública de Render quedará accesible a quien la conozca; restrinja acceso a nivel de red/plataforma si maneja información institucional. Sin una barrera externa no se puede considerar segura frente a visitantes no autorizados.

## Separación estricta
- Incidencias: tabla `incidencias` en Supabase.
- Urgencias: hojas actuales `Ordenes_Urgencias` y `Sin_Orden_Urgencias` de `BASE_URGENCIAS` en Google Sheets.
- Evidencias: carpetas e IDs de Google Drive existentes, sin reubicación.
- La migración histórica está deshabilitada; no hay eliminación de registros entre módulos.
- Se mantienen los identificadores originales y el flujo de captura existente.

## Seguridad y pruebas
- No incluya `.env`, claves, tokens o JSON de credenciales en ZIP/Git.
- Rotar claves si se compartieron fuera de un entorno autorizado.
- Probar en servicio de staging con permisos mínimos antes de producción.
- Las operaciones automáticas de actualización diaria de seguimiento permanecen activas: revise su comportamiento en staging.
- Esta entrega separa el núcleo y la interfaz y aplica protecciones puntuales; **no** implementa autenticación ni garantiza aislamiento frente a usuarios de una URL pública.
- No se ejecutaron pruebas de integración con los servicios reales.
