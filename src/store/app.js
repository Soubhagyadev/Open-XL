/**
 * Minimal reactive store using the Observer pattern.
 * No external dependency needed.
 */

const state = {
  isLoading: false,
  error: null,
  data: null,
};

const listeners = new Set();

export function getState() {
  return { ...state };
}

export function setState(partial) {
  Object.assign(state, partial);
  listeners.forEach((fn) => fn(getState()));
}

export function subscribe(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

/**
 * Mount the root UI into a DOM element.
 * Replace this with a framework (React/Vue) if needed.
 * @param {HTMLElement} container
 */
export function renderApp(container) {
  container.innerHTML = `<h2>Open-XL</h2><p>Add-in loaded successfully.</p>`;
}
