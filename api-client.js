/* Contrato compartido: solo acepta respuestas del backend actualizado. */
(function () {
  'use strict';
  const API_VERSION = '2026-09-05.2';
  const DEFAULT_URL = 'https://script.google.com/macros/s/AKfycbxwsWsvtHKWj_lrkmqXZLMVBNuXKGcZEb2yW5dPY1jzeK8ZNEqNn0XjDRc_WZNycnSeqA/exec';
  window.VeloraAPI = {
    url() { return localStorage.getItem('velora_sheets_url') || DEFAULT_URL; },
    newId() {
      if (typeof crypto.randomUUID === 'function') return 'VL-' + crypto.randomUUID();
      const bytes = crypto.getRandomValues(new Uint8Array(16));
      return 'VL-' + Array.from(bytes, value => value.toString(16).padStart(2, '0')).join('');
    },
    async request(action, data = {}, url = this.url()) {
      if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url)) {
        throw new Error('Configura una URL válida de Google Apps Script (/exec).');
      }
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 60000);
      try {
        const response = await fetch(url, {
          method: 'POST', redirect: 'follow', credentials: 'omit',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ ...data, action }), signal: controller.signal
        });
        const text = await response.text();
        let result;
        try { result = JSON.parse(text); } catch (_) {
          throw new Error('El servidor no devolvió JSON. Revisa los permisos y la implementación de Apps Script.');
        }
        if (!response.ok || result.api_version !== API_VERSION) {
          throw new Error('El backend desplegado está desactualizado. Publica la nueva versión de Google Apps Script.');
        }
        if (result.status !== 'success') {
          const error = new Error(result.message || 'El servidor rechazó la operación.');
          error.code = result.code;
          throw error;
        }
        return result;
      } catch (error) {
        if (error.name === 'AbortError') throw new Error('No se pudo confirmar la respuesta. Reintenta: se conservará el mismo pedido.');
        throw error;
      } finally { clearTimeout(timeout); }
    }
  };
})();
