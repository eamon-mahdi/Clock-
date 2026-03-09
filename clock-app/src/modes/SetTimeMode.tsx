import React, { useState, useCallback } from 'react';
import type { TimeValue } from '../types';
import AnalogClock from '../components/AnalogClock';
import DigitalDisplay from '../components/DigitalDisplay';
import { randomTime, timesMatch, formatTime, timeLabel } from '../utils/time';
import { playSuccess, playError, playClick, playTick } from '../utils/sound';

type Difficulty = 'easy' | 'medium' | 'hard';
type RoundState = 'playing' | 'correct' | 'wrong';

const TOTAL_ROUNDS = 8;

const SetTimeMode: React.FC = () => {
  const [difficulty, setDifficulty] = useState<Difficulty | null>(null);
  const [target, setTarget] = useState<TimeValue>({ hours: 3, minutes: 0 });
  const [current, setCurrent] = useState<TimeValue>({ hours: 12, minutes: 0 });
  const [roundState, setRoundState] = useState<RoundState>('playing');
  const [score, setScore] = useState(0);
  const [round, setRound] = useState(1);
  const [finished, setFinished] = useState(false);
  const [highlightHour, setHighlightHour] = useState(false);
  const [highlightMinute, setHighlightMinute] = useState(false);

  const generateTarget = (diff: Difficulty): TimeValue => {
    return randomTime(diff !== 'hard');
  };

  const startGame = (diff: Difficulty) => {
    const t = generateTarget(diff);
    setDifficulty(diff);
    setTarget(t);
    setCurrent({ hours: 12, minutes: 0 });
    setRoundState('playing');
    setScore(0);
    setRound(1);
    setFinished(false);
    playClick();
  };

  const handleTimeChange = useCallback((t: TimeValue) => {
    setCurrent(t);
  }, []);

  const tolerance = difficulty === 'easy' ? 0 : difficulty === 'medium' ? 2 : 1;

  const handleCheck = () => {
    if (roundState !== 'playing') return;
    const correct = timesMatch(current, target, tolerance) &&
      (difficulty === 'easy' ? current.hours === target.hours : true);

    if (correct) {
      playSuccess();
      setScore(s => s + 1);
      setRoundState('correct');
    } else {
      playError();
      setRoundState('wrong');
    }
  };

  const handleNext = () => {
    if (!difficulty) return;
    playClick();
    if (round >= TOTAL_ROUNDS) {
      setFinished(true);
      return;
    }
    const t = generateTarget(difficulty);
    setTarget(t);
    setCurrent({ hours: 12, minutes: 0 });
    setRoundState('playing');
    setRound(r => r + 1);
  };

  const handleHint = (which: 'hour' | 'minute') => {
    if (which === 'hour') {
      setCurrent(c => ({ ...c, hours: target.hours }));
      setHighlightHour(true);
      setTimeout(() => setHighlightHour(false), 1200);
    } else {
      setCurrent(c => ({ ...c, minutes: target.minutes }));
      setHighlightMinute(true);
      setTimeout(() => setHighlightMinute(false), 1200);
    }
    playTick();
  };

  if (!difficulty) {
    return (
      <div className="mode-container">
        <h2 className="mode-title">⏰ Set the Clock</h2>
        <p className="mode-subtitle">Drag the hands to match the digital time shown!</p>
        <div className="difficulty-grid">
          {([
            { id: 'easy', label: '⭐ Easy', desc: 'Hours only' },
            { id: 'medium', label: '⭐⭐ Medium', desc: '5-minute steps' },
            { id: 'hard', label: '⭐⭐⭐ Hard', desc: 'Any minute!' },
          ] as { id: Difficulty; label: string; desc: string }[]).map(d => (
            <button key={d.id} className={`difficulty-btn ${d.id}`} onClick={() => startGame(d.id)}>
              <span className="diff-label">{d.label}</span>
              <span className="diff-desc">{d.desc}</span>
            </button>
          ))}
        </div>
      </div>
    );
  }

  if (finished) {
    const pct = Math.round((score / TOTAL_ROUNDS) * 100);
    const emoji = pct >= 90 ? '🏆' : pct >= 70 ? '🌟' : pct >= 50 ? '👍' : '💪';
    return (
      <div className="mode-container">
        <div className="result-screen">
          <div className="result-emoji">{emoji}</div>
          <h2 className="result-title">Well Done!</h2>
          <div className="result-score">{score} / {TOTAL_ROUNDS}</div>
          <div className="result-pct">{pct}%</div>
          <p className="result-msg">
            {pct >= 90 ? 'You\'re a clock-setting superstar!' :
             pct >= 70 ? 'Really good work!' :
             pct >= 50 ? 'Good try! Keep practising!' :
             'Practice makes perfect!'}
          </p>
          <div style={{ display: 'flex', gap: 12, justifyContent: 'center', marginTop: 20 }}>
            <button className="btn-primary" onClick={() => startGame(difficulty!)}>Play Again</button>
            <button className="btn-secondary" onClick={() => setDifficulty(null)}>Change Level</button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="mode-container">
      <div className="quiz-header">
        <div className="quiz-stat">
          <span className="quiz-stat-label">Round</span>
          <span className="quiz-stat-value">{round}/{TOTAL_ROUNDS}</span>
        </div>
        <div className="quiz-stat">
          <span className="quiz-stat-label">Score</span>
          <span className="quiz-stat-value">{score}</span>
        </div>
      </div>

      <div className="progress-bar-wrap">
        <div className="progress-bar-fill" style={{ width: `${((round - 1) / TOTAL_ROUNDS) * 100}%` }} />
      </div>

      <p style={{ textAlign: 'center', fontWeight: 700, fontSize: 18, color: '#5c3d1e', marginBottom: 16 }}>
        Set the clock hands to show this time:
      </p>

      <div style={{ textAlign: 'center', marginBottom: 20 }}>
        <DigitalDisplay time={target} showLabel large />
      </div>

      <AnalogClock
        time={current}
        interactive={roundState === 'playing'}
        showLabels
        showMinuteTicks
        highlightHour={highlightHour}
        highlightMinute={highlightMinute}
        onTimeChange={handleTimeChange}
        size={280}
      />

      <div style={{ textAlign: 'center', marginTop: 8, fontSize: 14, color: '#888' }}>
        Your clock shows: <strong style={{ color: '#5c3d1e' }}>{formatTime(current)}</strong>
      </div>

      {roundState === 'playing' && (
        <div style={{ textAlign: 'center', marginTop: 16 }}>
          <button className="btn-primary btn-large" onClick={handleCheck}>
            ✅ Check My Answer
          </button>
          <div style={{ marginTop: 12, display: 'flex', gap: 10, justifyContent: 'center' }}>
            <button className="hint-btn small" onClick={() => handleHint('hour')}>
              💡 Hint: Hour hand
            </button>
            <button className="hint-btn small" onClick={() => handleHint('minute')}>
              💡 Hint: Minute hand
            </button>
          </div>
        </div>
      )}

      {roundState !== 'playing' && (
        <div style={{ textAlign: 'center', marginTop: 16 }}>
          <div className={`feedback-banner ${roundState}`}>
            {roundState === 'correct'
              ? `✅ Correct! That's ${formatTime(target)} — ${timeLabel(target)}`
              : `❌ Not quite! The answer was ${formatTime(target)} — ${timeLabel(target)}`}
          </div>
          {roundState === 'wrong' && (
            <AnalogClock time={target} interactive={false} showLabels size={200} />
          )}
          <button className="btn-primary" style={{ marginTop: 14 }} onClick={handleNext}>
            {round >= TOTAL_ROUNDS ? 'See Results 🎉' : 'Next Round →'}
          </button>
        </div>
      )}
    </div>
  );
};

export default SetTimeMode;
