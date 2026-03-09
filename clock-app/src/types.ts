export type GameMode = 'menu' | 'learn' | 'quiz' | 'set-time';

export interface TimeValue {
  hours: number;   // 0-11
  minutes: number; // 0-59
}

export interface QuizResult {
  correct: boolean;
  targetTime: TimeValue;
  givenTime: TimeValue;
}
