# Activación de Pixel y Purchase

El código local está preparado. La URL de Apps Script actualmente publicada ejecuta una versión anterior. Actualizar primero el backend y después el sitio. El frontend nuevo rechaza respuestas antiguas para evitar confirmar pedidos que no se guardaron.

## 1. Google Apps Script

En el proyecto vinculado a la hoja de pedidos:

1. Reemplazar el código por google_apps_script.js.
2. En Configuración > Propiedades del Script, configurar:
   - tokem: token autorizado para Meta CAPI. El backend también acepta META_CAPI_ACCESS_TOKEN como alias.
   - META_PIXEL_ID: 1635217208321987.
   - ADMIN_SECRET: contraseña privada de al menos 16 caracteres; será la contraseña de acceso a admin.html. No incluirla en archivos públicos.
   - META_TEST_EVENT_CODE: vacío para producción.
3. Ejecutar setupSheet en el editor y conceder los permisos de Sheets y peticiones externas que solicite Google.
4. Implementar > Administrar implementaciones > editar la implementación actual > Nueva versión > Implementar. Mantener la URL /exec actual. Ejecutar como propietario; acceso público al endpoint, con las acciones administrativas protegidas por contraseña validada en servidor.
5. Abrir /exec. Debe devolver api_version: "2026-09-05.2" y datos del servicio, sin listar clientes.

Guardar el archivo sin actualizar la implementación no actualiza /exec. Apps Script no carga los archivos .env de este repositorio.

Si se crea otra implementación con una URL distinta, cambiar la URL predeterminada en api-client.js, index.html y admin.html. La opción del panel solo modifica la configuración de ese navegador; no cambia la URL para los demás compradores.

## 2. Publicar la web

Publicar juntos index.html, admin.html y api-client.js en el alojamiento actual. Mantener las carpetas de imágenes y videos. No subir .env, .git, tests ni credenciales al alojamiento público.

## 3. Verificación real

1. En Meta Events Manager, seleccionar el dataset 1635217208321987.
2. Abrir la landing desde Eventos de prueba. Verificar PageView, ViewContent, InitiateCheckout y, al enviar un pedido de prueba, Lead.
3. Confirmar que el pedido aparece en el panel tras iniciar sesión con ADMIN_SECRET. La pantalla de éxito y WhatsApp solo continúan tras confirmar el guardado.
4. Para probar CAPI sin registrar una venta real, configurar el código de Eventos de prueba y usar un pedido dedicado de prueba. Marcar ese pedido como pagado en el panel. Debe aparecer Solo prueba y Meta debe mostrar Purchase desde servidor.
5. No reintentar ese pedido ficticio en producción. Mantenerlo identificado como prueba y excluirlo de las métricas comerciales, o retirarlo manualmente de la hoja tras la validación.
6. Vaciar META_TEST_EVENT_CODE. Confirmar después un pago real verificado en Yape/Plin. El panel muestra CAPI OK únicamente si Meta devuelve events_received: 1.

Las pruebas de código usan servicios simulados y no generan pedidos ni eventos reales:

    node --test tests/integration.test.cjs

## Comportamiento de Purchase

- Registrar un pedido envía Lead en navegador; Purchase se envía tras verificar el pago desde el panel.
- La venta se cierra por WhatsApp, por lo que action_source es chat.
- La llamada de servidor usa Graph API v26.0.
- El servidor usa el importe del catálogo y PEN, con la atribución del comprador guardada al registrar el pedido.
- Teléfono y nombre se normalizan y se envían con SHA-256. No se envía DNI a Meta.
- Cada compra usa purchase_<ID del pedido>. Los reintentos conservan ese ID y la fecha original; una compra aceptada no vuelve a enviarse.
- Un fallo de Meta conserva el pago confirmado y permite reintentar. El detalle queda en la nota de la celda Estado CAPI.
- Las fechas nuevas se guardan en ISO con zona horaria. Las antiguas sin zona se interpretan como hora de Lima. No se alteran fechas antiguas para eludir el límite temporal de Meta.
- La aceptación de la API no garantiza atribución a un anuncio. Esa comprobación se hace en Events Manager y depende de la coincidencia y atribución de Meta.
