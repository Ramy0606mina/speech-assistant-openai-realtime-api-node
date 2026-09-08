// Only the owner's new message can authorize a personal reminder. Forwarded
// material remains available for details, but cannot authorize a calendar write.
export function directOwnerRequestText(email) {
  const raw = String(email?.body?.content || email?.bodyPreview || '');
  const top = raw.split(/<blockquote\b|<div\b[^>]*\bid=["']divRplyFwdMsg["']|<hr\b/i)[0]
    .replace(/<(?:br|\/p|\/div)\b[^>]*>/gi, '\n').replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;|&#160;/gi, ' ').replace(/&amp;/gi, '&')
    .split(/(?:^|\n)\s*(?:From:|De\s*:|On .+wrote:|Le .+écrit\s*:|[-_]{3,}|Begin forwarded message:)/i)[0];
  return top.trim();
}

export function ownerReminderRequest(email) {
  const top = directOwnerRequestText(email);
  if (/\b(?:do not|don't|never|cancel|remove|delete)\b/i.test(top)) return '';
  return /\b(?:remind\s+me|(?:set|create|add|schedule|put)\b[^\n.!?]{0,120}\b(?:reminder|calendar))\b/i.test(top) ? top.trim() : '';
}
