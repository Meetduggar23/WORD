import React from 'react';
import { describe, it, expect, beforeEach } from 'vitest';
import { render, screen, fireEvent, act } from '@testing-library/react';
import { ThemeProvider, useTheme, Theme } from '../hooks/useTheme';

/* ── Controllable matchMedia mock (jsdom's is static) ──────────────── */
interface FakeMql {
  matches: boolean;
  listeners: Array<(e: { matches: boolean }) => void>;
}
let darkQuery: FakeMql;

beforeEach(() => {
  darkQuery = { matches: false, listeners: [] };
  Object.defineProperty(window, 'matchMedia', {
    writable: true,
    value: (query: string) => ({
      matches: darkQuery.matches,
      media: query,
      onchange: null,
      addListener: () => undefined,
      removeListener: () => undefined,
      addEventListener: (_: string, cb: (e: { matches: boolean }) => void) => { darkQuery.listeners.push(cb); },
      removeEventListener: (_: string, cb: (e: { matches: boolean }) => void) => {
        darkQuery.listeners = darkQuery.listeners.filter((l) => l !== cb);
      },
      dispatchEvent: () => false,
    }),
  });
  localStorage.clear();
  document.documentElement.removeAttribute('data-theme');
});

/* Probe component exposing the context */
const Probe: React.FC = () => {
  const { theme, resolvedTheme, setTheme, toggleTheme } = useTheme();
  return (
    <div>
      <span data-testid="theme">{theme}</span>
      <span data-testid="resolved">{resolvedTheme}</span>
      <button data-testid="set-light" onClick={() => setTheme('light' as Theme)}>light</button>
      <button data-testid="set-dark" onClick={() => setTheme('dark' as Theme)}>dark</button>
      <button data-testid="set-system" onClick={() => setTheme('system' as Theme)}>system</button>
      <button data-testid="toggle" onClick={toggleTheme}>toggle</button>
    </div>
  );
};

const renderApp = () => render(<ThemeProvider><Probe /></ThemeProvider>);

describe('ThemeProvider — light / dark / system', () => {
  it('defaults to system (resolved light when OS is light)', () => {
    renderApp();
    expect(screen.getByTestId('theme').textContent).toBe('system');
    expect(screen.getByTestId('resolved').textContent).toBe('light');
    expect(document.documentElement.getAttribute('data-theme')).toBe('light');
  });

  it('system mode resolves dark when the OS prefers dark', () => {
    darkQuery.matches = true;
    renderApp();
    expect(screen.getByTestId('resolved').textContent).toBe('dark');
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
    expect(document.documentElement.style.colorScheme).toBe('dark');
  });

  it('switches to explicit dark and light on demand', () => {
    renderApp();
    fireEvent.click(screen.getByTestId('set-dark'));
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
    expect(document.documentElement.style.colorScheme).toBe('dark');
    expect(localStorage.getItem('word.theme')).toBe('dark');

    fireEvent.click(screen.getByTestId('set-light'));
    expect(document.documentElement.getAttribute('data-theme')).toBe('light');
    expect(document.documentElement.style.colorScheme).toBe('light');
    expect(localStorage.getItem('word.theme')).toBe('light');
  });

  it('reacts live to OS changes while in system mode', () => {
    darkQuery.matches = false;
    renderApp();
    expect(document.documentElement.getAttribute('data-theme')).toBe('light');

    act(() => {
      darkQuery.matches = true;
      darkQuery.listeners.forEach((l) => l({ matches: true }));
    });
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
  });

  it('does not react to OS changes once an explicit theme is chosen', () => {
    renderApp();
    fireEvent.click(screen.getByTestId('set-dark'));

    act(() => {
      darkQuery.matches = false;
      darkQuery.listeners.forEach((l) => l({ matches: false }));
    });
    // Still dark — explicit choice wins over OS
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
  });

  it('restores the persisted theme from localStorage', () => {
    localStorage.setItem('word.theme', 'dark');
    renderApp();
    expect(screen.getByTestId('theme').textContent).toBe('dark');
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
  });

  it('falls back to system on a corrupted stored value', () => {
    localStorage.setItem('word.theme', 'sepia');
    renderApp();
    expect(screen.getByTestId('theme').textContent).toBe('system');
  });

  it('toggleTheme flips between light and dark', () => {
    renderApp();
    fireEvent.click(screen.getByTestId('set-light'));
    fireEvent.click(screen.getByTestId('toggle'));
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark');
    fireEvent.click(screen.getByTestId('toggle'));
    expect(document.documentElement.getAttribute('data-theme')).toBe('light');
  });

  it('syncs the browser theme-color meta to the active theme', () => {
    // jsdom doesn't load index.html — recreate the head tag it ships with
    document.head.insertAdjacentHTML('beforeend', '<meta name="theme-color" content="#1b3a5c" />');
    renderApp();
    fireEvent.click(screen.getByTestId('set-dark'));
    expect(document.querySelector('meta[name="theme-color"]')?.getAttribute('content')).toBe('#122a4f');
    fireEvent.click(screen.getByTestId('set-light'));
    expect(document.querySelector('meta[name="theme-color"]')?.getAttribute('content')).toBe('#103f91');
  });
});
