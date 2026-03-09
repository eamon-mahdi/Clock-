import type { TimeValue } from '../types';

export function formatTime(t: TimeValue): string {
  const h = t.hours === 0 ? 12 : t.hours;
  const m = t.minutes.toString().padStart(2, '0');
  return `${h}:${m}`;
}

export function timeLabel(t: TimeValue): string {
  const h = t.hours === 0 ? 12 : t.hours;
  if (t.minutes === 0) return `${h} o'clock`;
  if (t.minutes === 15) return `quarter past ${h}`;
  if (t.minutes === 30) return `half past ${h}`;
  if (t.minutes === 45) {
    const nextH = t.hours === 11 ? 12 : t.hours + 1;
    return `quarter to ${nextH}`;
  }
  if (t.minutes < 30) return `${t.minutes} past ${h}`;
  const nextH = t.hours === 11 ? 12 : t.hours + 1;
  return `${60 - t.minutes} to ${nextH}`;
}

export function randomTime(fiveMinOnly = false): TimeValue {
  const hours = Math.floor(Math.random() * 12);
  const minutes = fiveMinOnly
    ? Math.floor(Math.random() * 12) * 5
    : Math.floor(Math.random() * 60);
  return { hours, minutes };
}

export function timesMatch(a: TimeValue, b: TimeValue, tolerance = 2): boolean {
  const aMins = a.hours * 60 + a.minutes;
  const bMins = b.hours * 60 + b.minutes;
  const diff = Math.abs(aMins - bMins);
  return Math.min(diff, 720 - diff) <= tolerance;
}

export function anglesToTime(hourAngle: number, minuteAngle: number): TimeValue {
  // angles in degrees (0 = 12 o'clock, clockwise)
  const minutes = Math.round(((minuteAngle % 360) / 360) * 60) % 60;
  const hours = Math.round(((hourAngle % 360) / 360) * 12) % 12;
  return { hours, minutes };
}

export function timeToAngles(t: TimeValue): { hourAngle: number; minuteAngle: number } {
  const minuteAngle = (t.minutes / 60) * 360;
  const hourAngle = ((t.hours % 12) / 12) * 360 + (t.minutes / 60) * 30;
  return { hourAngle, minuteAngle };
}
