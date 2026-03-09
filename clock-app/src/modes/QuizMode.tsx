import React, { useState, useCallback } from 'react';
import type { TimeValue } from '../types';
import AnalogClock from '../components/AnalogClock';
import { randomTime, formatTime, timeLabel } from '../utils/time';
import { playSuccess, playError, playClick } from '../utils/sound';

type Difficulty = 'easy' | 'medium' | 'hard';
type AnswerState = 'pending' | 'correct' | 'wrong';

interface QuizQuestion {
  target: TimeValue;
  options: string[];
  correctIndex: number;
}

function generateOptions(target: TimeValue, difficulty: Difficulty): { options: string[]; correctIndex: number } {
  const correct = formatTime(target);
  const distractors: string[] = [];

  while (distractors.length < 3) {
    let t: TimeValue;
    if (difficulty === 'easy') {
      // Hour only variations
      const hours = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0].filter(h => h !== target.hours);
      t = { hours: hours[Math.floor(Math.random() * hours.length)], minutes: 0 };
    } else if (difficulty === 'medium') {
      // 5-minute increments
      const minsOpts = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55].filter(m => m !== target.minutes);
      t = { hours: target.hours, minutes: minsOpts[Math.floor(Math.random() * minsOpts.length)] };
      if (Math.random() > 0.5) t.hours = (target.hours + 1 + Math.floor(Math.random() * 10)) % 12;
    } else {
      t = randomTime(false);
    }
    const f = formatTime(t);
    if (f !== correct && !distractors.includes(f)) {
      distractors.push(f);
    }
  }

  const all = [correct, ...distractors].sort(() => Math.random() - 0.5);
  return { options: all, correctIndex: all.indexOf(correct) };
}

function buildQuestion(difficulty: Difficulty): QuizQuestion {
  const target = randomTime(difficulty !== 'hard');
  const { options, correctIndex } = generateOptions(target, difficulty);
  return { target, options, correctIndex };
}

const STREAK_GOAL = 5;
const TOTAL_QUESTIONS = 10;

const QuizMode: React.FC = () => {
  const [difficulty, setDifficulty] = useState<Difficulty | null>(null);
  const [question, setQuestion] = useState<QuizQuestion | null>(null);
  const [answerState, setAnswerState] = useState<AnswerState>('pending');
  const [score, setScore] = useState(0);
  const [streak, setStreak] = useState(0);
  const [questionNum, setQuestionNum] = useState(1);
  const [finished, setFinished] = useState(false);
  const [selectedIdx, setSelectedIdx] = useState<number | null>(null);

  const newQuestion = useCallback((diff: Difficulty) => {
    setQuestion(buildQuestion(diff));
    setAnswerState('pending');
    setSelectedIdx(null);
  }, []);

  const startGame = (diff: Difficulty) => {
    setDifficulty(diff);
    setScore(0);
    setStreak(0);
    setQuestionNum(1);
    setFinished(false);
    setQuestion(buildQuestion(diff));
    setAnswerState('pending');
    setSelectedIdx(null);
    playClick();
  };

  const handleAnswer = (idx: number) => {
    if (answerState !== 'pending' || !question) return;
    setSelectedIdx(idx);
    const correct = idx === question.correctIndex;
    setAnswerState(correct ? 'correct' : 'wrong');
    if (correct) {
      playSuccess();
      setScore(s => s + 1);
      setStreak(s => s + 1);
    } else {
      playError();
      setStreak(0);
    }
  };

  const handleNext = () => {
    if (!difficulty) return;
    playClick();
    if (questionNum >= TOTAL_QUESTIONS) {
      setFinished(true);
    } else {
      setQuestionNum(n => n + 1);
      newQuestion(difficulty);
    }
  };

  const handleRestart = () => {
    setDifficulty(null);
    setFinished(false);
  };

  // Difficulty picker
  if (!difficulty) {
    return (
      <div className="mode-container">
        <h2 className="mode-title">❓ Read the Clock Quiz</h2>
        <p className="mode-subtitle">Look at the clock and choose the correct time!</p>
        <div className="difficulty-grid">
          {([
            { id: 'easy', label: '⭐ Easy', desc: 'Hours only, 12 choices' },
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
    const pct = Math.round((score / TOTAL_QUESTIONS) * 100);
    const emoji = pct >= 90 ? '🏆' : pct >= 70 ? '🌟' : pct >= 50 ? '👍' : '💪';
    return (
      <div className="mode-container">
        <div className="result-screen">
          <div className="result-emoji">{emoji}</div>
          <h2 className="result-title">Quiz Complete!</h2>
          <div className="result-score">{score} / {TOTAL_QUESTIONS}</div>
          <div className="result-pct">{pct}%</div>
          <p className="result-msg">
            {pct >= 90 ? 'Amazing! You\'re a clock master!' :
             pct >= 70 ? 'Great job! Keep practising!' :
             pct >= 50 ? 'Good effort! Try again!' :
             'Keep practising, you\'ll get it!'}
          </p>
          <div style={{ display: 'flex', gap: 12, justifyContent: 'center', marginTop: 20 }}>
            <button className="btn-primary" onClick={() => startGame(difficulty!)}>Play Again</button>
            <button className="btn-secondary" onClick={handleRestart}>Change Level</button>
          </div>
        </div>
      </div>
    );
  }

  if (!question) return null;

  return (
    <div className="mode-container">
      <div className="quiz-header">
        <div className="quiz-stat">
          <span className="quiz-stat-label">Question</span>
          <span className="quiz-stat-value">{questionNum}/{TOTAL_QUESTIONS}</span>
        </div>
        <div className="quiz-stat">
          <span className="quiz-stat-label">Score</span>
          <span className="quiz-stat-value">{score}</span>
        </div>
        <div className="quiz-stat">
          <span className="quiz-stat-label">Streak</span>
          <span className="quiz-stat-value">{streak >= STREAK_GOAL ? '🔥' : streak > 0 ? `🔥x${streak}` : '—'}</span>
        </div>
      </div>

      <div className="progress-bar-wrap">
        <div className="progress-bar-fill" style={{ width: `${((questionNum - 1) / TOTAL_QUESTIONS) * 100}%` }} />
      </div>

      <p style={{ textAlign: 'center', fontWeight: 700, fontSize: 18, marginBottom: 12, color: '#5c3d1e' }}>
        What time does this clock show?
      </p>

      <AnalogClock time={question.target} interactive={false} showLabels size={260} />

      {answerState !== 'pending' && (
        <div style={{ textAlign: 'center', marginTop: 8, fontSize: 15, color: '#666' }}>
          ({timeLabel(question.target)})
        </div>
      )}

      <div className="options-grid">
        {question.options.map((opt, i) => {
          let cls = 'option-btn';
          if (answerState !== 'pending') {
            if (i === question.correctIndex) cls += ' option-correct';
            else if (i === selectedIdx) cls += ' option-wrong';
            else cls += ' option-dim';
          }
          return (
            <button
              key={i}
              className={cls}
              onClick={() => handleAnswer(i)}
              disabled={answerState !== 'pending'}
            >
              {opt}
              {answerState !== 'pending' && i === question.correctIndex && ' ✓'}
              {answerState !== 'pending' && i === selectedIdx && i !== question.correctIndex && ' ✗'}
            </button>
          );
        })}
      </div>

      {answerState !== 'pending' && (
        <div style={{ textAlign: 'center', marginTop: 16 }}>
          <div className={`feedback-banner ${answerState}`}>
            {answerState === 'correct' ? `✅ Correct! It's ${formatTime(question.target)}` : `❌ The answer was ${formatTime(question.target)}`}
          </div>
          <button className="btn-primary" style={{ marginTop: 14 }} onClick={handleNext}>
            {questionNum >= TOTAL_QUESTIONS ? 'See Results 🎉' : 'Next Question →'}
          </button>
        </div>
      )}
    </div>
  );
};

export default QuizMode;
