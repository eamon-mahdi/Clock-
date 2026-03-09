import React, { useState } from 'react';
import type { TimeValue } from '../types';
import AnalogClock from '../components/AnalogClock';
import DigitalDisplay from '../components/DigitalDisplay';
import { playClick } from '../utils/sound';
import { timeLabel } from '../utils/time';

const LearnMode: React.FC = () => {
  const [time, setTime] = useState<TimeValue>({ hours: 3, minutes: 0 });
  const [showHint, setShowHint] = useState(false);

  const handleTimeChange = (t: TimeValue) => setTime(t);

  const presets = [
    { label: "3:00", time: { hours: 3, minutes: 0 } },
    { label: "6:30", time: { hours: 6, minutes: 30 } },
    { label: "9:15", time: { hours: 9, minutes: 15 } },
    { label: "12:45", time: { hours: 0, minutes: 45 } },
    { label: "7:20", time: { hours: 7, minutes: 20 } },
    { label: "11:55", time: { hours: 11, minutes: 55 } },
  ];

  return (
    <div className="mode-container">
      <h2 className="mode-title">
        🎓 Learn the Clock
      </h2>
      <p className="mode-subtitle">
        Drag the <span style={{ color: '#c0392b', fontWeight: 700 }}>red hour hand</span> and{' '}
        <span style={{ color: '#3a5a9c', fontWeight: 700 }}>blue minute hand</span> to set the time!
      </p>

      <div className="clock-and-display">
        <AnalogClock
          time={time}
          interactive
          showLabels
          showMinuteTicks
          onTimeChange={handleTimeChange}
          size={300}
        />
        <div style={{ marginTop: 24 }}>
          <DigitalDisplay time={time} showLabel large />
        </div>
      </div>

      {/* Hint section */}
      <div style={{ marginTop: 20, textAlign: 'center' }}>
        <button
          className="hint-btn"
          onClick={() => { setShowHint(!showHint); playClick(); }}
        >
          {showHint ? '🙈 Hide Hint' : '💡 How to read the clock'}
        </button>

        {showHint && (
          <div className="hint-box">
            <div className="hint-item">
              <span className="hand-dot hour-dot" /> The <strong>short red hand</strong> points to the <strong>hour</strong>
            </div>
            <div className="hint-item">
              <span className="hand-dot minute-dot" /> The <strong>long blue hand</strong> points to the <strong>minutes</strong>
            </div>
            <div className="hint-item">
              📍 When the blue hand points to <strong>12</strong>, it's exactly the hour (0 minutes)
            </div>
            <div className="hint-item">
              📍 When the blue hand points to <strong>3</strong>, it's 15 minutes (quarter past)
            </div>
            <div className="hint-item">
              📍 When the blue hand points to <strong>6</strong>, it's 30 minutes (half past)
            </div>
            <div className="hint-item">
              📍 When the blue hand points to <strong>9</strong>, it's 45 minutes (quarter to)
            </div>
          </div>
        )}
      </div>

      {/* Preset times */}
      <div style={{ marginTop: 24, textAlign: 'center' }}>
        <p style={{ color: '#666', marginBottom: 10, fontSize: 15, fontWeight: 600 }}>Try these times:</p>
        <div className="preset-grid">
          {presets.map((p) => (
            <button
              key={p.label}
              className="preset-btn"
              onClick={() => { setTime(p.time); playClick(); }}
            >
              {p.label}
              <span className="preset-label">{timeLabel(p.time)}</span>
            </button>
          ))}
        </div>
      </div>

      {/* Minute reference */}
      <div style={{ marginTop: 28 }}>
        <MinuteReference />
      </div>
    </div>
  );
};

const MinuteReference: React.FC = () => {
  const items = [
    { num: 1, mins: 5 }, { num: 2, mins: 10 }, { num: 3, mins: 15 },
    { num: 4, mins: 20 }, { num: 5, mins: 25 }, { num: 6, mins: 30 },
    { num: 7, mins: 35 }, { num: 8, mins: 40 }, { num: 9, mins: 45 },
    { num: 10, mins: 50 }, { num: 11, mins: 55 }, { num: 12, mins: 0 },
  ];
  return (
    <div className="minute-ref">
      <h3 style={{ marginBottom: 12, color: '#5c3d1e', fontSize: 16 }}>
        📌 Minute Hand Reference
      </h3>
      <div className="minute-ref-grid">
        {items.map(({ num, mins }) => (
          <div key={num} className="minute-ref-item">
            <span className="minute-ref-num">{num}</span>
            <span className="minute-ref-arrow">→</span>
            <span className="minute-ref-mins">{mins === 0 ? '00' : mins} min</span>
          </div>
        ))}
      </div>
    </div>
  );
};

export default LearnMode;
