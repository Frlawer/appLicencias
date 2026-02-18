function getAdminWhitelist_() {
  const raw = getConfigValue_('ADMIN_ALLOWED_EMAILS', '');
  if (!raw) return [];
  return raw
    .split(',')
    .map(email => String(email).trim().toLowerCase())
    .filter(Boolean);
}

function esAdminAutorizado_() {
  const userEmail = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!userEmail) return false;

  const whitelist = getAdminWhitelist_();
  if (!whitelist.length) {
    return true;
  }

  return whitelist.includes(userEmail);
}

function assertAdminAutorizado_() {
  if (!esAdminAutorizado_()) {
    throw new Error('Acceso denegado: usuario no autorizado para operación administrativa');
  }
}
