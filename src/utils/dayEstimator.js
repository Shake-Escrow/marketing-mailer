/**
 * Given consecutive days of email counts, estimates what day the
 * window starts on by minimizing SSE against y ≈ a*x² + c.
 *
 * @param {number[]} emails - Array of observed email counts
 * @param {number} a - Quadratic coefficient (default 0.246)
 * @param {number} c - Constant term (default 1.75)
 * @param {number} currentDayOffset - Offset from the first observed day to the current day
 * @returns {{ startDay: number, currentDay: number, sse: number }}
 */
export function findCurrentDay(emails, a = 0.246, c = 1.75, currentDayOffset = emails.length - 1) {
  if (!Array.isArray(emails) || emails.length < 1) {
    throw new Error("Expected at least one day of data");
  }

  function sse(d) {
    return emails.reduce((sum, y, i) => {
      const predicted = a * (d + i) ** 2 + c;
      return sum + (y - predicted) ** 2;
    }, 0);
  }

  // Coarse grid search over plausible day range
  let bestD = 1;
  let bestSSE = Infinity;

  for (let d = 1; d <= 1000; d += 0.1) {
    const s = sse(d);
    if (s < bestSSE) {
      bestSSE = s;
      bestD = d;
    }
  }

  // Refine with a narrow golden-section search around the coarse best
  let lo = bestD - 1, hi = bestD + 1;
  const phi = (Math.sqrt(5) - 1) / 2;

  for (let iter = 0; iter < 100; iter++) {
    const m1 = hi - phi * (hi - lo);
    const m2 = lo + phi * (hi - lo);
    if (sse(m1) < sse(m2)) {
      hi = m2;
    } else {
      lo = m1;
    }
  }

  const refinedD = (lo + hi) / 2;
  const startDay = Math.round(refinedD);
  const currentDay = startDay + currentDayOffset;

  return {
    startDay,
    currentDay,
    sse: sse(startDay),
  };
}
