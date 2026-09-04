export function triggerHaptic(ms = 40) {
  if (typeof window !== 'undefined' && window.navigator && window.navigator.vibrate) {
    try {
      window.navigator.vibrate(ms);
    } catch {
      /* not all browsers support vibrate; harmless no-op */
    }
  }
}

export function copyTextSafely(text) {
  if (typeof navigator !== 'undefined' && navigator.clipboard && window.isSecureContext) {
    return navigator.clipboard.writeText(text);
  }
  return new Promise((resolve, reject) => {
    try {
      const el = document.createElement('textarea');
      el.value = text;
      el.style.position = 'fixed';
      el.style.left = '-999999px';
      el.style.top = '-999999px';
      document.body.appendChild(el);
      el.focus();
      el.select();
      document.execCommand('copy');
      document.body.removeChild(el);
      resolve();
    } catch (err) {
      reject(err);
    }
  });
}

export function playArrivalChime(enabled) {
  if (!enabled) return;
  try {
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    if (ctx.state === 'suspended') ctx.resume();

    const now = ctx.currentTime;
    const osc1 = ctx.createOscillator();
    const g1 = ctx.createGain();
    osc1.type = 'sine';
    osc1.frequency.setValueAtTime(659.25, now);
    g1.gain.setValueAtTime(0.25, now);
    g1.gain.exponentialRampToValueAtTime(0.0001, now + 0.45);
    osc1.connect(g1);
    g1.connect(ctx.destination);
    osc1.start(now);
    osc1.stop(now + 0.45);

    const osc2 = ctx.createOscillator();
    const g2 = ctx.createGain();
    osc2.type = 'triangle';
    osc2.frequency.setValueAtTime(880.0, now + 0.1);
    g2.gain.setValueAtTime(0.3, now + 0.1);
    g2.gain.exponentialRampToValueAtTime(0.0001, now + 0.85);
    osc2.connect(g2);
    g2.connect(ctx.destination);
    osc2.start(now + 0.1);
    osc2.stop(now + 0.85);
  } catch {
    /* audio can fail (autoplay policy, no AudioContext) — non-critical */
  }
}
