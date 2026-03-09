import React from 'react';
import type { TimeValue } from '../types';
import { formatTime, timeLabel } from '../utils/time';

interface Props {
  time: TimeValue;
  showLabel?: boolean;
  large?: boolean;
}

const DigitalDisplay: React.FC<Props> = ({ time, showLabel = true, large = false }) => {
  return (
    <div style={{
      textAlign: 'center',
      background: 'linear-gradient(135deg, #1a1a2e, #16213e)',
      borderRadius: 16,
      padding: large ? '18px 32px' : '12px 24px',
      display: 'inline-block',
      boxShadow: '0 4px 20px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.1)',
      border: '2px solid #0f3460',
    }}>
      <div style={{
        fontFamily: "'Courier New', 'Digital-7', monospace",
        fontSize: large ? 56 : 38,
        fontWeight: 'bold',
        color: '#00e676',
        textShadow: '0 0 20px rgba(0,230,118,0.6), 0 0 40px rgba(0,230,118,0.3)',
        letterSpacing: 4,
        lineHeight: 1,
      }}>
        {formatTime(time)}
      </div>
      {showLabel && (
        <div style={{
          marginTop: 6,
          fontSize: large ? 18 : 14,
          color: '#80cbc4',
          fontFamily: "'Nunito', sans-serif",
          fontWeight: 600,
          letterSpacing: 1,
        }}>
          {timeLabel(time)}
        </div>
      )}
    </div>
  );
};

export default DigitalDisplay;
