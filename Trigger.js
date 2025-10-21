function onEdit(e) {
  try {
    const ss      = e.source;
    const sh      = ss.getActiveSheet();
    const name    = sh.getName();
    const fila    = e.range.getRow();
    const col     = e.range.getColumn();
    const valor   = String(e.value || '').trim().toUpperCase();

    if (name !== 'bbdd reservas horas trabajar' && name !== 'bbdd reservas horas librar') {
      Logger.log('Hoja distinta, no intervengo: ' + name);
      return;
    }
    if (fila <= 1) {
      Logger.log('Fila ≤1, ignoro.');
      return;
    }

    // Detectar columnas relevantes
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idxVal  = headers.indexOf('Validación') + 1;
    if (!idxVal) {
      Logger.log("No encuentro columna 'Validación'");
      return;
    }
    const idxCorreo = headers.indexOf('Correo') + 1;
    const idxCamp   = headers.indexOf('Campaña') + 1;
    const idxFecha  = headers.indexOf('Fecha') + 1;
    const idxHoras  = headers.indexOf('HORAS') + 1;
    const idxFran   = headers.indexOf('FRANJA') + 1;
    const idxTipo   = headers.indexOf('Tipo') + 1;

    // Leer datos de la fila editada
    const datos = sh.getRange(fila, 1, 1, sh.getLastColumn()).getValues()[0];
    const campaña = datos[idxCamp-1];
    const fecha   = datos[idxFecha-1];
    const horas   = datos[idxHoras-1];
    const franja  = datos[idxFran-1];
    const correo  = datos[idxCorreo-1];
    const validacion = datos[idxVal-1];
    let tipoReserva = '';

    // Tipo según columna "Tipo"
    if (name === 'bbdd reservas horas trabajar') {
      tipoReserva = (datos[idxTipo-1] || 'TRABAJAR').toString().toUpperCase();
    } else if (name === 'bbdd reservas horas librar') {
      tipoReserva = 'LIBRAR';
    }

    // Validar email
    if (!correo || correo.toString().indexOf('@') < 0) {
      Logger.log('Email inválido o vacío, no envío.');
      return;
    }

    Logger.log('Editado columna: ' + col + ', idxVal: ' + idxVal + ', valor: "' + valor + '"');

    // Detectar si la edición es en columna "Validación"
    if (col === idxVal) {
      Logger.log('Columna validación editada, valor: "' + valor + '"');
      // Solo enviar email si valor es KO
      if (valor && valor.toUpperCase().trim() === 'KO') {
        // Enviar email de cancelación
        const fechaFmt = Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        let cuerpoHTML = `
          <div style="font-family:Segoe UI,Arial,sans-serif; font-size:1.13em;">
            <p>¡Hola!</p>
            <p>
              Tu solicitud ha sido cancelada:
            </p>
            <ul style="line-height:1.7;">
              <li><b>Campaña:</b> ${campaña}</li>
              <li><b>Fecha:</b> ${fechaFmt}</li>
              <li><b>Horario:</b> ${franja}</li>
              <li><b>Horas:</b> ${horas}</li>
            </ul>
            <p>
              Si tienes dudas, contacta con tu coordinador/a.<br>Equipo de solicitudes
            </p>
          </div>
        `;
        const tipoHoja = (tipoReserva === 'LIBRAR') ? ' (Librar)' : (tipoReserva === 'COBRAR') ? ' (Cobrar)' : '';
        const asunto = `Solicitud cancelada – Bolsa de horas${tipoHoja}`;

        MailApp.sendEmail({
          to: correo,
          subject: asunto,
          htmlBody: cuerpoHTML
        });
        Logger.log('Email KO enviado a ' + correo + ' (' + asunto + ')');
      } else {
        Logger.log('Valor Validación no es KO, no envío email.');
      }
      return;
    }

    // No es edición en columna Validación: detectar nueva reserva
    // Condiciones: Campaña, Fecha, Horas, Franja y Correo no vacíos y Validación vacía
    const validacionStr = (validacion || '').toString().trim();
    if (
      campaña && campaña.toString().trim() !== '' &&
      fecha && fecha.toString().trim() !== '' &&
      horas && horas.toString().trim() !== '' &&
      franja && franja.toString().trim() !== '' &&
      correo && correo.toString().indexOf('@') >= 0 &&
      validacionStr === ''
    ) {
      // Detectar si antes no había correo para esta fila
      // Para eso, obtenemos el valor anterior de la celda editada si es posible
      // Pero e.oldValue no siempre está definido, así que también permitimos enviar si se completa toda la fila con correo válido y validación vacía

      // Solo enviar email si la edición fue en columna Correo o en alguna columna clave que complete la reserva
      // Para simplificar, enviamos email si la fila cumple condiciones y la edición no fue en Validación (ya descartado)

      const fechaFmt = Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
      let cuerpoHTML = `
        <div style="font-family:Segoe UI,Arial,sans-serif; font-size:1.13em;">
          <p>¡Hola!</p>
          <p>
            Tu solicitud ha sido registrada correctamente (<b>${tipoReserva}</b>):
          </p>
          <ul style="line-height:1.7;">
            <li><b>Campaña:</b> ${campaña}</li>
            <li><b>Fecha:</b> ${fechaFmt}</li>
            <li><b>Horario:</b> ${franja}</li>
            <li><b>Horas:</b> ${horas}</li>
          </ul>
          <p>
            ¡Gracias!<br>Equipo de solicitudes
          </p>
        </div>
      `;
      const tipoHoja = (tipoReserva === 'LIBRAR') ? ' (Librar)' : (tipoReserva === 'COBRAR') ? ' (Cobrar)' : '';
      const asunto = `Solicitud recibida – Bolsa de horas${tipoHoja}`;

      MailApp.sendEmail({
        to: correo,
        subject: asunto,
        htmlBody: cuerpoHTML
      });
      Logger.log('Email registro enviado a ' + correo + ' (' + asunto + ')');
    } else {
      Logger.log('No se cumplen condiciones para enviar email de registro.');
    }

  } catch (err) {
    Logger.log('Error en onEdit: ' + err);
  }
}
