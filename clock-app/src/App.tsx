import React, { useState } from 'react';
import type { GameMode } from './types';
import LearnMode from './modes/LearnMode';
import QuizMode from './modes/QuizMode';
import SetTimeMode from './modes/SetTimeMode';
import { playClick } from './utils/sound';
import './App.css';

const App: React.FC = () => {
  const [mode, setMode] = useState<GameMode>('menu');

  const goHome = () => { setMode('menu'); playClick(); };

  return (
    <div className="app-root">
      <header className="app-header">
        <div className="header-inner">
          <button className="home-btn" onClick={goHome} title="Back to menu">
            🏠
          </button>
          <h1 className="app-title">
            <span className="title-clock">🕐</span>
            Clock Learning Fun!
          </h1>
          <div style={{ width: 44 }} />
        </div>
      </header>

      <main className="app-main">
        {mode === 'menu' && <MainMenu onSelect={setMode} />}
        {mode === 'learn' && <LearnMode />}
        {mode === 'quiz' && <QuizMode />}
        {mode === 'set-time' && <SetTimeMode />}
      </main>

      <footer className="app-footer">
        Made with ❤️ to help you learn to tell time!
      </footer>
    </div>
  );
};

interface MainMenuProps {
  onSelect: (mode: GameMode) => void;
}

const MainMenu: React.FC<MainMenuProps> = ({ onSelect }) => {
  const cards = [
    {
      id: 'learn' as GameMode,
      emoji: '📚',
      title: 'Learn',
      subtitle: 'Explore the clock',
      desc: 'Drag the clock hands, see the time, and learn how to read a clock!',
      color: '#2e7d32',
      bg: 'linear-gradient(135deg, #e8f5e9, #c8e6c9)',
      border: '#81c784',
    },
    {
      id: 'quiz' as GameMode,
      emoji: '❓',
      title: 'Quiz Me!',
      subtitle: 'Read the clock',
      desc: 'Look at the clock and pick the correct time. Can you get them all right?',
      color: '#e65100',
      bg: 'linear-gradient(135deg, #fff3e0, #ffe0b2)',
      border: '#ffb74d',
    },
    {
      id: 'set-time' as GameMode,
      emoji: '⏰',
      title: 'Set the Time',
      subtitle: 'Match the clock',
      desc: 'Drag the hands to match the digital time. Be as accurate as you can!',
      color: '#6a1b9a',
      bg: 'linear-gradient(135deg, #f3e5f5, #e1bee7)',
      border: '#ce93d8',
    },
  ];

  // Decorative clock SVG
  const hourAngle = (300 - 90) * (Math.PI / 180);
  const minAngle = (60 - 90) * (Math.PI / 180);

  return (
    <div className="menu-container">
      <div className="menu-hero">
        <svg width="110" height="110" viewBox="0 0 120 120" className="menu-clock-svg">
          <circle cx="60" cy="60" r="55" fill="#fffef8" stroke="#d4a847" strokeWidth="5" />
          {[1,2,3,4,5,6,7,8,9,10,11,12].map(n => {
            const a = ((n / 12) * 360 - 90) * (Math.PI / 180);
            return (
              <text
                key={n}
                x={60 + 42 * Math.cos(a)}
                y={60 + 42 * Math.sin(a)}
                textAnchor="middle"
                dominantBaseline="central"
                fontSize="11"
                fontWeight="700"
                fill="#5c3d1e"
                fontFamily="Nunito, sans-serif"
              >{n}</text>
            );
          })}
          <line x1="60" y1="60" x2={60 + 28 * Math.cos(hourAngle)} y2={60 + 28 * Math.sin(hourAngle)} stroke="#c0392b" strokeWidth="5" strokeLinecap="round" />
          <line x1="60" y1="60" x2={60 + 40 * Math.cos(minAngle)} y2={60 + 40 * Math.sin(minAngle)} stroke="#3a5a9c" strokeWidth="3.5" strokeLinecap="round" />
          <circle cx="60" cy="60" r="5" fill="#5c3d1e" />
        </svg>
        <div className="menu-hero-text">
          <h2 className="menu-headline">Hello! Let's learn to tell the time! 🎉</h2>
          <p className="menu-subhead">Choose an activity below to get started:</p>
        </div>
      </div>

      <div className="menu-cards">
        {cards.map(c => (
          <button
            key={c.id}
            className="menu-card"
            style={{ background: c.bg, borderColor: c.border }}
            onClick={() => { onSelect(c.id); playClick(); }}
          >
            <div className="card-emoji">{c.emoji}</div>
            <div className="card-title" style={{ color: c.color }}>{c.title}</div>
            <div className="card-subtitle">{c.subtitle}</div>
            <div className="card-desc">{c.desc}</div>
            <div className="card-arrow" style={{ color: c.color }}>Start →</div>
          </button>
        ))}
      </div>

      <div className="menu-tip">
        <strong>💡 Tip:</strong> Start with <em>Learn</em> if you're new to clocks!
      </div>
    </div>
  );
};

export default App;
