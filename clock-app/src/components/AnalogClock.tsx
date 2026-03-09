import React, { useCallback, useRef, useState, useEffect } from 'react';
import type { TimeValue } from '../types';
import { timeToAngles } from '../utils/time';
import { playTick } from '../utils/sound';

interface Props {
  time: TimeValue;
  interactive?: boolean;
  showLabels?: boolean;
  showMinuteTicks?: boolean;
  highlightHour?: boolean;
  highlightMinute?: boolean;
  onTimeChange?: (t: TimeValue) => void;
  size?: number;
}

type Dragging = 'hour' | 'minute' | null;

const AnalogClock: React.FC<Props> = ({
  time,
  interactive = false,
  showLabels = true,
  showMinuteTicks = true,
  highlightHour = false,
  highlightMinute = false,
  onTimeChange,
  size = 320,
}) => {
  const svgRef = useRef<SVGSVGElement>(null);
  const dragging = useRef<Dragging>(null);
  const lastMinute = useRef(time.minutes);
  const [hovering, setHovering] = useState<'hour' | 'minute' | null>(null);

  const { hourAngle, minuteAngle } = timeToAngles(time);
  const cx = size / 2;
  const cy = size / 2;
  const r = size / 2 - 10;

  const getAngleFromCenter = useCallback((clientX: number, clientY: number) => {
    if (!svgRef.current) return 0;
    const rect = svgRef.current.getBoundingClientRect();
    const x = clientX - rect.left - cx;
    const y = clientY - rect.top - cy;
    let angle = Math.atan2(y, x) * (180 / Math.PI) + 90;
    if (angle < 0) angle += 360;
    return angle;
  }, [cx, cy]);

  const updateFromAngle = useCallback((angle: number, which: Dragging) => {
    if (!which || !onTimeChange) return;
    if (which === 'minute') {
      const newMinutes = Math.round((angle / 360) * 60) % 60;
      const newHourAngle = ((time.hours % 12) / 12) * 360 + (newMinutes / 60) * 30;
      if (newMinutes !== lastMinute.current) {
        playTick();
        lastMinute.current = newMinutes;
      }
      onTimeChange({ hours: time.hours, minutes: newMinutes });
      return { hourAngle: newHourAngle, minuteAngle: (newMinutes / 60) * 360 };
    } else {
      const rawHours = (angle / 360) * 12;
      const newHours = Math.round(rawHours) % 12;
      onTimeChange({ hours: newHours, minutes: time.minutes });
    }
  }, [time, onTimeChange]);

  useEffect(() => {
    const onMove = (e: MouseEvent | TouchEvent) => {
      if (!dragging.current) return;
      const clientX = 'touches' in e ? e.touches[0].clientX : e.clientX;
      const clientY = 'touches' in e ? e.touches[0].clientY : e.clientY;
      const angle = getAngleFromCenter(clientX, clientY);
      updateFromAngle(angle, dragging.current);
    };
    const onUp = () => { dragging.current = null; };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    window.addEventListener('touchmove', onMove, { passive: true });
    window.addEventListener('touchend', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      window.removeEventListener('touchmove', onMove);
      window.removeEventListener('touchend', onUp);
    };
  }, [getAngleFromCenter, updateFromAngle]);

  const handProps = (which: 'hour' | 'minute') => ({
    style: { cursor: interactive ? 'grab' : 'default' },
    onMouseDown: interactive ? (e: React.MouseEvent) => {
      e.preventDefault();
      dragging.current = which;
    } : undefined,
    onTouchStart: interactive ? (_e: React.TouchEvent) => {
      dragging.current = which;
    } : undefined,
    onMouseEnter: interactive ? () => setHovering(which) : undefined,
    onMouseLeave: interactive ? () => setHovering(null) : undefined,
  });

  // Build number positions
  const numbers = Array.from({ length: 12 }, (_, i) => {
    const n = i + 1;
    const angle = ((n / 12) * 360 - 90) * (Math.PI / 180);
    const nr = r - (showLabels ? 28 : 22);
    return {
      n,
      x: cx + nr * Math.cos(angle),
      y: cy + nr * Math.sin(angle),
    };
  });

  // Minute ticks
  const minuteTicks = Array.from({ length: 60 }, (_, i) => {
    const angle = (i / 60) * 2 * Math.PI - Math.PI / 2;
    const isHour = i % 5 === 0;
    const outer = r - 2;
    const inner = outer - (isHour ? 14 : 7);
    return {
      x1: cx + outer * Math.cos(angle),
      y1: cy + outer * Math.sin(angle),
      x2: cx + inner * Math.cos(angle),
      y2: cy + inner * Math.sin(angle),
      isHour,
    };
  });

  // Hand endpoints
  const hourLen = r * 0.55;
  const minuteLen = r * 0.78;
  const hourRad = (hourAngle - 90) * (Math.PI / 180);
  const minRad = (minuteAngle - 90) * (Math.PI / 180);

  const hourTip = { x: cx + hourLen * Math.cos(hourRad), y: cy + hourLen * Math.sin(hourRad) };
  const minTip = { x: cx + minuteLen * Math.cos(minRad), y: cy + minuteLen * Math.sin(minRad) };

  // Back of hands
  const hourBack = { x: cx - (hourLen * 0.18) * Math.cos(hourRad), y: cy - (hourLen * 0.18) * Math.sin(hourRad) };
  const minBack = { x: cx - (minuteLen * 0.15) * Math.cos(minRad), y: cy - (minuteLen * 0.15) * Math.sin(minRad) };

  const isHourHovered = hovering === 'hour' || highlightHour;
  const isMinHovered = hovering === 'minute' || highlightMinute;

  return (
    <svg
      ref={svgRef}
      width={size}
      height={size}
      viewBox={`0 0 ${size} ${size}`}
      style={{ userSelect: 'none', touchAction: 'none', display: 'block', margin: '0 auto' }}
    >
      {/* Drop shadow */}
      <defs>
        <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="4" stdDeviation="8" floodOpacity="0.15" />
        </filter>
        <radialGradient id="clockFace" cx="40%" cy="35%">
          <stop offset="0%" stopColor="#fffef8" />
          <stop offset="100%" stopColor="#f0ece0" />
        </radialGradient>
      </defs>

      {/* Clock face */}
      <circle cx={cx} cy={cy} r={r} fill="url(#clockFace)" stroke="#d4a847" strokeWidth="6" filter="url(#shadow)" />
      <circle cx={cx} cy={cy} r={r - 8} fill="none" stroke="#f0d080" strokeWidth="2" opacity="0.5" />

      {/* Minute ticks */}
      {showMinuteTicks && minuteTicks.map((t, i) => (
        <line
          key={i}
          x1={t.x1} y1={t.y1}
          x2={t.x2} y2={t.y2}
          stroke={t.isHour ? '#8b5e3c' : '#c4a882'}
          strokeWidth={t.isHour ? 2.5 : 1.2}
          strokeLinecap="round"
        />
      ))}

      {/* Numbers */}
      {numbers.map(({ n, x, y }) => (
        <text
          key={n}
          x={x} y={y}
          textAnchor="middle"
          dominantBaseline="central"
          fontSize={size * 0.072}
          fontWeight="700"
          fontFamily="'Nunito', 'Comic Sans MS', cursive"
          fill="#5c3d1e"
        >
          {n}
        </text>
      ))}

      {/* Minute hand */}
      <line
        x1={minBack.x} y1={minBack.y}
        x2={minTip.x} y2={minTip.y}
        stroke={isMinHovered ? '#2196f3' : '#3a5a9c'}
        strokeWidth={isMinHovered ? size * 0.028 : size * 0.022}
        strokeLinecap="round"
        filter={isMinHovered ? "url(#shadow)" : undefined}
        {...handProps('minute')}
      />

      {/* Hour hand */}
      <line
        x1={hourBack.x} y1={hourBack.y}
        x2={hourTip.x} y2={hourTip.y}
        stroke={isHourHovered ? '#f44336' : '#c0392b'}
        strokeWidth={isHourHovered ? size * 0.042 : size * 0.034}
        strokeLinecap="round"
        filter={isHourHovered ? "url(#shadow)" : undefined}
        {...handProps('hour')}
      />

      {/* Center cap */}
      <circle cx={cx} cy={cy} r={size * 0.028} fill="#5c3d1e" />
      <circle cx={cx} cy={cy} r={size * 0.014} fill="#f0d080" />

      {/* Interactive hint circles at tips */}
      {interactive && (
        <>
          <circle
            cx={minTip.x} cy={minTip.y} r={size * 0.05}
            fill={isMinHovered ? 'rgba(33,150,243,0.25)' : 'rgba(33,150,243,0.08)'}
            {...handProps('minute')}
          />
          <circle
            cx={hourTip.x} cy={hourTip.y} r={size * 0.055}
            fill={isHourHovered ? 'rgba(244,67,54,0.25)' : 'rgba(244,67,54,0.08)'}
            {...handProps('hour')}
          />
        </>
      )}
    </svg>
  );
};

export default AnalogClock;
