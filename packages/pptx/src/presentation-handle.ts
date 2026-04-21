import type { MediaElement, Slide } from './types';
import { drawPlayBadge } from './media-chrome';

const EMU_PER_PX = 9525;
const emuToPx = (v: number, scale: number) => (v / EMU_PER_PX) * scale;

interface MediaState {
  el: MediaElement;
  /** Interaction bounds — drives hit testing and progress bar placement. For
   *  audio this is enlarged below the icon so the bar has room. */
  rect: { x: number; y: number; w: number; h: number };
  /** The original poster/icon bounds from EMU coords. Badge anchors to this
   *  so the play/pause circle stays on top of the icon, not on the bar. */
  posterRect: { x: number; y: number; w: number; h: number };
  media: HTMLVideoElement | HTMLAudioElement;
  objectUrl: string;
}

export interface PresentationHandle {
  play(mediaPath?: string): void;
  pause(mediaPath?: string): void;
  dispose(): void;
}

export interface PresentOptions {
  /** Display width in CSS pixels (same value used to render the base slide). */
  width: number;
  /** Device pixel ratio used when sizing the canvas backing store. */
  dpr: number;
  /** Slide width in EMU (from Presentation.slideWidth). */
  slideWidthEmu: number;
  /** Retrieve the raw media bytes for the given zip path. */
  fetchMedia: (mediaPath: string) => Promise<Blob>;
  /** Draw the static slide (poster + play badge) onto the canvas. */
  drawBase: () => Promise<void>;
}

/**
 * Attach canvas-native playback to a slide. Layers each media element's video
 * frame (via HTMLVideoElement → drawImage) and a self-drawn play/progress
 * overlay on top of the base rendering. Click on a media element to toggle
 * playback. Call dispose() to stop the RAF loop and release blob URLs.
 */
export async function createPresentationHandle(
  canvas: HTMLCanvasElement,
  slide: Slide,
  opts: PresentOptions,
): Promise<PresentationHandle> {
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('2D context not available');

  const slideScale = opts.width / (opts.slideWidthEmu / EMU_PER_PX);

  await opts.drawBase();

  const base = document.createElement('canvas');
  base.width = canvas.width;
  base.height = canvas.height;
  const baseCtx = base.getContext('2d');
  if (!baseCtx) throw new Error('base 2D context not available');
  baseCtx.drawImage(canvas, 0, 0);

  const mediaElements = slide.elements.filter((e): e is MediaElement => e.type === 'media');
  const states: MediaState[] = [];

  for (const el of mediaElements) {
    let blob: Blob;
    try {
      blob = await opts.fetchMedia(el.mediaPath);
    } catch {
      continue;
    }
    const typed = new Blob([blob], { type: el.mimeType || blob.type });
    const url = URL.createObjectURL(typed);
    const media: HTMLVideoElement | HTMLAudioElement = el.mediaKind === 'video'
      ? document.createElement('video')
      : document.createElement('audio');
    media.src = url;
    media.preload = 'metadata';
    if (el.mediaKind === 'video') {
      (media as HTMLVideoElement).playsInline = true;
    }
    const posterRect = {
      x: emuToPx(el.x, slideScale),
      y: emuToPx(el.y, slideScale),
      w: emuToPx(el.width, slideScale),
      h: emuToPx(el.height, slideScale),
    };
    // Audio icons in pptx are often small (40–80 px); a bar that narrow is
    // unusable, so expand the control surface horizontally and push it
    // downward below the icon so the bar sits clearly under the speaker —
    // not across its bottom edge. Video keeps its native bounds since the
    // bar overlays the frame.
    const rect = el.mediaKind === 'audio'
      ? {
          x: posterRect.x + posterRect.w / 2 - Math.max(posterRect.w, 260) / 2,
          y: posterRect.y,
          w: Math.max(posterRect.w, 260),
          h: posterRect.h + 36,
        }
      : posterRect;
    states.push({ el, rect, posterRect, media, objectUrl: url });
  }

  let rafId: number | null = null;
  let disposed = false;
  let hoveredState: MediaState | null = null;

  const drawFrame = () => {
    ctx.setTransform(opts.dpr, 0, 0, opts.dpr, 0, 0);
    const cssW = canvas.width / opts.dpr;
    const cssH = canvas.height / opts.dpr;
    ctx.drawImage(base, 0, 0, canvas.width, canvas.height, 0, 0, cssW, cssH);

    for (const s of states) {
      const media = s.media;

      // Draw the current video frame whenever the element has pixel data —
      // paused or playing. Paused should hold the frame you stopped on, not
      // snap back to the poster (which is the base render underneath).
      if (s.el.mediaKind === 'video' && media.readyState >= 2) {
        const { x, y, w, h } = s.posterRect;
        ctx.drawImage(media as HTMLVideoElement, x, y, w, h);
      }

      // Controls fade in on hover. While actively dragging the scrubber
      // (`seeking`), keep the hovered state latched to that element so the
      // bar doesn't disappear when the pointer strays off-element mid-drag.
      const isActive = s === hoveredState || seeking?.state === s;
      if (isActive) drawControls(ctx, s, media);
    }
  };

  const tick = () => {
    if (disposed) return;
    drawFrame();
    rafId = requestAnimationFrame(tick);
  };

  const toLocal = (clientX: number, clientY: number) => {
    const bb = canvas.getBoundingClientRect();
    const cssW = canvas.width / opts.dpr;
    const cssH = canvas.height / opts.dpr;
    return {
      x: ((clientX - bb.left) / bb.width) * cssW,
      y: ((clientY - bb.top) / bb.height) * cssH,
    };
  };

  type Hit =
    | { kind: 'seek'; state: MediaState; fraction: number }
    | { kind: 'toggle'; state: MediaState };

  const hitTest = (localX: number, localY: number): Hit | null => {
    for (const s of states) {
      const { x, y, w, h } = s.rect;
      if (localX < x || localX > x + w || localY < y || localY > y + h) continue;
      const bar = barGeometry(s);
      // Expand the bar's vertical hit area so a 2–4 px visual bar is still easy
      // to grab on high-DPI displays; horizontally use the visual bounds.
      const hitYTop = bar.y - 12;
      const hitYBot = bar.y + bar.h + 8;
      const duration = Number.isFinite(s.media.duration) ? s.media.duration : 0;
      if (
        duration > 0 &&
        localX >= bar.x && localX <= bar.x + bar.w &&
        localY >= hitYTop && localY <= hitYBot
      ) {
        const fraction = Math.max(0, Math.min(1, (localX - bar.x) / bar.w));
        return { kind: 'seek', state: s, fraction };
      }
      return { kind: 'toggle', state: s };
    }
    return null;
  };

  let seeking: { state: MediaState; wasPlaying: boolean } | null = null;

  const seekTo = (state: MediaState, fraction: number) => {
    const duration = Number.isFinite(state.media.duration) ? state.media.duration : 0;
    if (duration <= 0) return;
    state.media.currentTime = duration * fraction;
  };

  const onPointerDown = (e: PointerEvent) => {
    const { x: lx, y: ly } = toLocal(e.clientX, e.clientY);
    const hit = hitTest(lx, ly);
    if (!hit) return;
    if (hit.kind === 'seek') {
      seeking = { state: hit.state, wasPlaying: !hit.state.media.paused };
      hit.state.media.pause();
      seekTo(hit.state, hit.fraction);
      canvas.setPointerCapture(e.pointerId);
      e.preventDefault();
    } else {
      if (hit.state.media.paused) void hit.state.media.play().catch(() => undefined);
      else hit.state.media.pause();
    }
  };

  const onPointerMove = (e: PointerEvent) => {
    const { x: lx, y: ly } = toLocal(e.clientX, e.clientY);

    // Hover tracking — update which element the pointer is over each move.
    hoveredState = null;
    for (const s of states) {
      const { x, y, w, h } = s.rect;
      if (lx >= x && lx <= x + w && ly >= y && ly <= y + h) {
        hoveredState = s;
        break;
      }
    }

    if (seeking) {
      const bar = barGeometry(seeking.state);
      const fraction = Math.max(0, Math.min(1, (lx - bar.x) / bar.w));
      seekTo(seeking.state, fraction);
    }
  };

  const onPointerLeave = () => { hoveredState = null; };

  const onPointerUp = (e: PointerEvent) => {
    if (!seeking) return;
    const { wasPlaying, state } = seeking;
    seeking = null;
    canvas.releasePointerCapture(e.pointerId);
    if (wasPlaying) void state.media.play().catch(() => undefined);
  };

  canvas.addEventListener('pointerdown', onPointerDown);
  canvas.addEventListener('pointermove', onPointerMove);
  canvas.addEventListener('pointerleave', onPointerLeave);
  canvas.addEventListener('pointerup', onPointerUp);
  canvas.addEventListener('pointercancel', onPointerUp);
  canvas.style.cursor = 'pointer';
  tick();

  return {
    play(mediaPath) {
      for (const s of states) {
        if (!mediaPath || s.el.mediaPath === mediaPath) {
          void s.media.play().catch(() => undefined);
        }
      }
    },
    pause(mediaPath) {
      for (const s of states) {
        if (!mediaPath || s.el.mediaPath === mediaPath) s.media.pause();
      }
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      if (rafId !== null) cancelAnimationFrame(rafId);
      canvas.removeEventListener('pointerdown', onPointerDown);
      canvas.removeEventListener('pointermove', onPointerMove);
      canvas.removeEventListener('pointerleave', onPointerLeave);
      canvas.removeEventListener('pointerup', onPointerUp);
      canvas.removeEventListener('pointercancel', onPointerUp);
      canvas.style.cursor = '';
      for (const s of states) {
        s.media.pause();
        s.media.removeAttribute('src');
        s.media.load();
        URL.revokeObjectURL(s.objectUrl);
      }
    },
  };
}

const AUDIO_PILL_H = 28;
const AUDIO_PILL_SIDE_PAD = 14;
const AUDIO_TIME_RESERVE = 72;
const AUDIO_TIME_GAP = 10;
const BAR_H = 3;

function drawControls(
  ctx: CanvasRenderingContext2D,
  state: MediaState,
  media: HTMLVideoElement | HTMLAudioElement,
): void {
  const duration = Number.isFinite(media.duration) ? media.duration : 0;
  const progress = duration > 0 ? Math.min(1, media.currentTime / duration) : 0;

  const poster = state.posterRect;
  drawPlayBadge(
    ctx,
    poster.x + poster.w / 2,
    poster.y + poster.h / 2,
    poster.w,
    poster.h,
    media.paused ? 'paused' : 'playing',
  );

  if (state.el.mediaKind === 'audio') {
    drawAudioPill(ctx, state, media, duration, progress);
  } else {
    drawVideoChrome(ctx, state, media, duration, progress);
  }
}

function drawVideoChrome(
  ctx: CanvasRenderingContext2D,
  state: MediaState,
  media: HTMLVideoElement | HTMLAudioElement,
  duration: number,
  progress: number,
): void {
  const { x, y, w, h } = state.rect;
  const chromeH = Math.max(28, Math.min(56, h * 0.22));
  const chromeY = y + h - chromeH;

  ctx.save();
  const grad = ctx.createLinearGradient(0, chromeY, 0, y + h);
  grad.addColorStop(0, 'rgba(0, 0, 0, 0)');
  grad.addColorStop(1, 'rgba(0, 0, 0, 0.55)');
  ctx.fillStyle = grad;
  ctx.fillRect(x, chromeY, w, chromeH);
  ctx.restore();

  const bar = barGeometry(state);
  drawBar(ctx, bar, progress, duration > 0);

  const fontPx = 11;
  ctx.save();
  ctx.font = `500 ${fontPx}px system-ui, -apple-system, sans-serif`;
  ctx.textBaseline = 'middle';
  ctx.shadowColor = 'rgba(0, 0, 0, 0.75)';
  ctx.shadowBlur = 3;
  ctx.fillStyle = 'rgba(255, 255, 255, 0.95)';
  drawFixedSlotTime(ctx, media.currentTime, duration, bar.x, bar.y - 10, 'bottom');
  ctx.restore();
}

function drawAudioPill(
  ctx: CanvasRenderingContext2D,
  state: MediaState,
  media: HTMLVideoElement | HTMLAudioElement,
  duration: number,
  progress: number,
): void {
  const stadium = audioStadium(state.rect);
  ctx.save();
  roundedRect(ctx, stadium.x, stadium.y, stadium.w, stadium.h, stadium.h / 2);
  ctx.fillStyle = 'rgba(20, 20, 20, 0.72)';
  ctx.fill();

  const fontPx = 11;
  ctx.font = `500 ${fontPx}px system-ui, -apple-system, sans-serif`;
  ctx.textBaseline = 'middle';
  ctx.fillStyle = 'rgba(255, 255, 255, 0.95)';
  drawFixedSlotTime(
    ctx, media.currentTime, duration,
    stadium.x + AUDIO_PILL_SIDE_PAD, stadium.y + stadium.h / 2, 'middle',
  );
  ctx.restore();

  const bar = barGeometry(state);
  drawBar(ctx, bar, progress, duration > 0);
}

/**
 * Render "m:ss / m:ss" such that the separator stays at a fixed x even as the
 * current-time digits tick, by right-aligning the current side into a fixed
 * slot and left-aligning the duration side. Side slot width is sized once per
 * call from the duration string so the layout fits longer clips too.
 */
function drawFixedSlotTime(
  ctx: CanvasRenderingContext2D,
  current: number,
  duration: number,
  x: number,
  anchorY: number,
  anchor: 'middle' | 'bottom',
): void {
  const curText = formatTime(current);
  const durText = formatTime(duration);
  const sep = ' / ';
  // Widest representation of either side: use the longer of the two formatted
  // strings. In practice duration dominates since current ≤ duration.
  const slotW = Math.max(ctx.measureText(curText).width, ctx.measureText(durText).width);
  const sepW = ctx.measureText(sep).width;
  const y = anchor === 'bottom' ? anchorY : anchorY;
  const prevAlign = ctx.textAlign;
  ctx.textAlign = 'right';
  ctx.fillText(curText, x + slotW, y);
  ctx.textAlign = 'left';
  ctx.fillText(sep, x + slotW, y);
  ctx.fillText(durText, x + slotW + sepW, y);
  ctx.textAlign = prevAlign;
}

function drawBar(
  ctx: CanvasRenderingContext2D,
  bar: { x: number; y: number; w: number; h: number },
  progress: number,
  hasDuration: boolean,
): void {
  const radius = bar.h / 2;
  ctx.save();
  roundedRect(ctx, bar.x, bar.y, bar.w, bar.h, radius);
  ctx.fillStyle = 'rgba(255, 255, 255, 0.35)';
  ctx.fill();

  if (progress > 0) {
    roundedRect(ctx, bar.x, bar.y, bar.w * progress, bar.h, radius);
    ctx.fillStyle = '#fff';
    ctx.fill();
  }

  if (hasDuration) {
    const handleR = 5;
    const handleX = Math.max(bar.x + handleR, Math.min(bar.x + bar.w - handleR, bar.x + bar.w * progress));
    ctx.shadowColor = 'rgba(0, 0, 0, 0.3)';
    ctx.shadowBlur = 3;
    ctx.fillStyle = '#fff';
    ctx.beginPath();
    ctx.arc(handleX, bar.y + bar.h / 2, handleR, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();
}

function audioStadium(rect: { x: number; y: number; w: number; h: number }) {
  const w = Math.max(220, rect.w - 24);
  return {
    x: rect.x + rect.w / 2 - w / 2,
    y: rect.y + rect.h - AUDIO_PILL_H - 4,
    w,
    h: AUDIO_PILL_H,
  };
}

function barGeometry(state: MediaState) {
  if (state.el.mediaKind === 'audio') {
    const stadium = audioStadium(state.rect);
    const barX = stadium.x + AUDIO_PILL_SIDE_PAD + AUDIO_TIME_RESERVE + AUDIO_TIME_GAP;
    const barW = Math.max(40, stadium.x + stadium.w - AUDIO_PILL_SIDE_PAD - barX);
    return {
      x: barX,
      y: stadium.y + (stadium.h - BAR_H) / 2,
      w: barW,
      h: BAR_H,
    };
  }
  const rect = state.rect;
  const margin = Math.max(12, rect.w * 0.025);
  const bottomPad = Math.max(12, Math.min(18, rect.h * 0.05));
  return {
    x: rect.x + margin,
    y: rect.y + rect.h - BAR_H - bottomPad,
    w: rect.w - margin * 2,
    h: BAR_H,
  };
}

function roundedRect(
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  r: number,
): void {
  const rr = Math.min(r, h / 2, w / 2);
  ctx.beginPath();
  ctx.moveTo(x + rr, y);
  ctx.lineTo(x + w - rr, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + rr);
  ctx.lineTo(x + w, y + h - rr);
  ctx.quadraticCurveTo(x + w, y + h, x + w - rr, y + h);
  ctx.lineTo(x + rr, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - rr);
  ctx.lineTo(x, y + rr);
  ctx.quadraticCurveTo(x, y, x + rr, y);
  ctx.closePath();
}

function formatTime(seconds: number): string {
  if (!Number.isFinite(seconds) || seconds < 0) return '0:00';
  const s = Math.floor(seconds);
  const m = Math.floor(s / 60);
  const ss = (s % 60).toString().padStart(2, '0');
  return `${m}:${ss}`;
}
