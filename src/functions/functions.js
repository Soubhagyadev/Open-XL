/**
 * Custom Excel functions (streaming / non-streaming).
 * Registered via CustomFunctions.associate().
 */

/**
 * Returns the current UTC timestamp.
 * @customfunction
 * @returns {string}
 */
export function NOW_UTC() {
  return new Date().toISOString();
}

CustomFunctions.associate("NOW_UTC", NOW_UTC);
