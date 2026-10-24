
/**
 * AGURIZA TRAINING — función de Netlify.
 *
 * Variables privadas obligatorias en Netlify:
 * APPS_SCRIPT_URL   = URL /exec de la nueva implementación de Apps Script.
 * APPS_SCRIPT_TOKEN = el mismo token guardado en las propiedades de Apps Script.
 */

exports.handler = async function (event) {
  if (event.httpMethod !== 'POST') {
    return json(405, {
      ok: false,
      error: 'Método no permitido.'
    });
  }

  const scriptUrl = process.env.APPS_SCRIPT_URL;
  const token = process.env.APPS_SCRIPT_TOKEN;

  if (!scriptUrl || !token) {
    console.error('Faltan APPS_SCRIPT_URL o APPS_SCRIPT_TOKEN.');
    return json(500, {
      ok: false,
      error: 'El servidor no está configurado. Revisa las variables de entorno de Netlify.'
    });
  }

  let request;

  try {
    request = JSON.parse(event.body || '{}');
  } catch (_) {
    return json(400, {
      ok: false,
      error: 'La solicitud no contiene JSON válido.'
    });
  }

  const allowedActions = new Set([
    'getPlanOptions',
    'saveSession'
  ]);

  if (!allowedActions.has(request.action)) {
    return json(400, {
      ok: false,
      error: 'Acción no válida.'
    });
  }

  const payload = {
    token: token,
    action: request.action,
    payload: request.payload || null
  };

  try {
    const upstream = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload),
      redirect: 'follow'
    });

    const responseText = await upstream.text();
    let result;

    try {
      result = JSON.parse(responseText);
    } catch (_) {
      console.error(
        'Apps Script no devolvió JSON. Estado:',
        upstream.status,
        'Respuesta:',
        responseText.slice(0, 300)
      );

      return json(502, {
        ok: false,
        error: 'Apps Script devolvió una respuesta inesperada. Verifica la URL /exec, los permisos y la implementación.'
      });
    }

    if (!upstream.ok) {
      console.error(
        'Respuesta HTTP no satisfactoria desde Apps Script:',
        upstream.status
      );

      return json(502, {
        ok: false,
        error: 'No fue posible comunicarse con Google Apps Script.'
      });
    }

    if (result.ok !== true) {
      return json(400, {
        ok: false,
        error: result.error || 'Apps Script rechazó la solicitud.'
      });
    }

    return json(200, {
      ok: true,
      data: result.data
    });

  } catch (error) {
    console.error(
      'Error conectando con Apps Script:',
      error && error.message ? error.message : error
    );

    return json(502, {
      ok: false,
      error: 'No se pudo conectar con Google Apps Script. Revisa la URL de implementación y vuelve a intentarlo.'
    });
  }
};

function json(statusCode, body) {
  return {
    statusCode: statusCode,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    },
    body: JSON.stringify(body)
  };
}
