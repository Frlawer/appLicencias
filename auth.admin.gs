function getAdminWhitelist_() {
  const raw = getConfigValue_('ADMIN_ALLOWED_EMAILS', '');
  if (!raw) return [];
  return raw
    .split(',')
    .map(email => String(email).trim().toLowerCase())
    .filter(Boolean);
}

function obtenerEmailSesionSeguro_() {
  // In public web app deployments, reading session email can throw auth-required errors.
  // We swallow those errors and return empty string to keep the app responsive.
  try {
    const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    if (activeEmail) return activeEmail;
  } catch (errorActive) {
    Logger.log('No se pudo leer Session.getActiveUser().getEmail(): ' + errorActive);
  }

  try {
    const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    if (effectiveEmail) return effectiveEmail;
  } catch (errorEffective) {
    Logger.log('No se pudo leer Session.getEffectiveUser().getEmail(): ' + errorEffective);
  }

  return '';
}

function esAdminAutorizado_() {
  const whitelist = getAdminWhitelist_();
  if (!whitelist.length) {
    return true;
  }

  const userEmail = obtenerEmailSesionSeguro_();
  if (!userEmail) return false;

  return whitelist.includes(userEmail);
}

function assertAdminAutorizado_() {
  if (!esAdminAutorizado_()) {
    throw new Error('Acceso denegado: usuario no autorizado para operación administrativa');
  }
}
