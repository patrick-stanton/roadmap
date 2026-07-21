/* =============================================================================
 * gantt-core.js
 * -----------------------------------------------------------------------------
 * Pure, DOM-free logic for the Capability Roadmap tool.
 *
 * WHY this file exists separately from the user interface (UI):
 *   The scheduling and CSV (Comma-Separated Values) logic is the part that must
 *   be correct — a wrong bar position or a mangled export is not recoverable
 *   once decisions have been made and pushed back to Cameo. Keeping it free of
 *   any browser / Document Object Model (DOM) dependency lets us run it under
 *   Node.js and assert its behaviour with real tests. The *exact same file* is
 *   then inlined into the single-file app, so what we test is what ships.
 *
 * Exposes a single object `GanttCore`. Works in the browser (as a global) and
 * under Node.js (via module.exports) with no build step.
 * ===========================================================================*/
(function (root) {
  'use strict';

  /* ---------------------------------------------------------------------------
   * SECTION 1 — CSV parsing & serialization
   * Cameo Systems Modeler (a Model-Based Systems Engineering [MBSE] tool)
   * exports tables to CSV. Real-world exports are messy: a leading Byte Order
   * Mark (BOM), quoted fields containing commas/newlines, doubled quotes as an
   * escape, and locale-dependent delimiters (comma vs. semicolon). We handle
   * all of these rather than assuming a clean file.
   * -------------------------------------------------------------------------*/

  /** Remove a leading UTF-8 BOM if present (Excel/Cameo frequently add one). */
  function stripBOM(text) {
    return text && text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  /** Return the substring up to the first line break (for delimiter sniffing). */
  function firstLine(text) {
    const i = text.search(/\r\n|\r|\n/);
    return i === -1 ? text : text.slice(0, i);
  }

  /**
   * Guess the field delimiter by counting candidate characters that sit OUTSIDE
   * quotes on the header line. We do this because we cannot know the locale of
   * the machine that produced the Cameo export.
   */
  function detectDelimiter(text) {
    const line = firstLine(stripBOM(text || ''));
    const counts = { ',': 0, ';': 0, '\t': 0, '|': 0 };
    let inQuotes = false;
    for (const ch of line) {
      if (ch === '"') inQuotes = !inQuotes;
      else if (!inQuotes && counts.hasOwnProperty(ch)) counts[ch]++;
    }
    let best = ',';
    let bestN = -1;
    for (const d in counts) {
      if (counts[d] > bestN) {
        bestN = counts[d];
        best = d;
      }
    }
    return best;
  }

  /**
   * Parse CSV text into an array of string arrays (rows of cells).
   * Implements RFC-4180-style quoting: fields may be wrapped in double quotes,
   * and a doubled quote ("") inside a quoted field is a literal quote.
   *
   * @param {string} text      Raw file contents.
   * @param {string} [delim]   Delimiter; auto-detected when omitted.
   * @returns {{rows: string[][], delimiter: string}}
   */
  function parseCSV(text, delim) {
    text = stripBOM(text || '');
    const delimiter = delim || detectDelimiter(text);
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    let i = 0;
    const n = text.length;

    while (i < n) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i++;
          continue;
        }
        field += ch;
        i++;
        continue;
      }
      // Not inside quotes:
      if (ch === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (ch === delimiter) {
        row.push(field);
        field = '';
        i++;
        continue;
      }
      if (ch === '\r') {
        i++;
        continue; // handled by the \n branch (CRLF) or ignored (lone CR)
      }
      if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        i++;
        continue;
      }
      field += ch;
      i++;
    }
    // Flush the final field/row (files may not end in a newline).
    row.push(field);
    rows.push(row);

    // Drop a trailing empty row produced when the file ends with a newline.
    if (
      rows.length &&
      rows[rows.length - 1].length === 1 &&
      rows[rows.length - 1][0] === ''
    ) {
      rows.pop();
    }
    return { rows: rows, delimiter: delimiter };
  }

  /**
   * Parse CSV into records keyed by header name.
   * Blank/duplicate headers are made safe so no column is silently lost.
   *
   * @returns {{headers: string[], records: Object[], delimiter: string}}
   */
  function parseTable(text, delim) {
    const { rows, delimiter } = parseCSV(text, delim);
    if (!rows.length) return { headers: [], records: [], delimiter: delimiter };
    const seen = {};
    const headers = rows[0].map(function (h, idx) {
      let name = (h || '').trim() || 'Column ' + (idx + 1);
      if (seen[name] != null) {
        seen[name]++;
        name = name + ' (' + seen[name] + ')';
      } else {
        seen[name] = 0;
      }
      return name;
    });
    const records = [];
    for (let r = 1; r < rows.length; r++) {
      const cells = rows[r];
      // Skip fully blank lines.
      if (cells.length === 1 && cells[0].trim() === '') continue;
      const obj = {};
      headers.forEach(function (h, c) {
        obj[h] = c < cells.length ? cells[c] : '';
      });
      records.push(obj);
    }
    return { headers: headers, records: records, delimiter: delimiter };
  }

  /** Quote a single cell only when required by the delimiter/quote/newline rules. */
  function escapeCell(value, delimiter) {
    const v = value == null ? '' : String(value);
    if (
      v.indexOf('"') !== -1 ||
      v.indexOf(delimiter) !== -1 ||
      v.indexOf('\n') !== -1 ||
      v.indexOf('\r') !== -1
    ) {
      return '"' + v.replace(/"/g, '""') + '"';
    }
    return v;
  }

  /**
   * Serialize an array of record objects back to CSV.
   * Uses CRLF line endings for maximum compatibility with Excel and Cameo.
   */
  function toCSV(headers, records, delim) {
    const delimiter = delim || ',';
    const lines = [
      headers
        .map(function (h) {
          return escapeCell(h, delimiter);
        })
        .join(delimiter),
    ];
    for (const rec of records) {
      lines.push(
        headers
          .map(function (h) {
            return escapeCell(rec[h], delimiter);
          })
          .join(delimiter)
      );
    }
    return lines.join('\r\n');
  }

  /* ---------------------------------------------------------------------------
   * SECTION 2 — value coercion helpers
   * -------------------------------------------------------------------------*/

  /** Parse a number, tolerating stray text/units; returns `fallback` on failure. */
  function parseNumber(value, fallback) {
    if (value == null || value === '') return fallback;
    const n = parseFloat(String(value).replace(/[^0-9.\-]/g, ''));
    return isFinite(n) ? n : fallback;
  }

  /**
   * Convert a free-text priority into a sortable number where LOWER = HIGHER
   * priority. Accepts numbers ("1", "2") or words ("High", "Must", "Low").
   */
  function priorityToRank(value) {
    if (value == null || value === '') return 999;
    const s = String(value).trim().toLowerCase();
    const words = {
      critical: 0,
      blocker: 0,
      highest: 1,
      must: 1,
      'must-have': 1,
      high: 2,
      should: 2,
      'should-have': 2,
      medium: 3,
      med: 3,
      normal: 3,
      could: 3,
      moderate: 3,
      low: 4,
      'nice-to-have': 5,
      lowest: 5,
      wont: 6,
      "won't": 6,
    };
    if (words[s] != null) return words[s];
    const n = parseFloat(s);
    return isFinite(n) ? n : 999;
  }

  /* ---------------------------------------------------------------------------
   * SECTION 3 — dates and week math
   * The roadmap's unit is the week. Everything on the timeline is expressed as
   * an integer "week index" relative to the project start (index 0). We anchor
   * weeks to Monday so bar edges land on clean week boundaries.
   * -------------------------------------------------------------------------*/

  const MS_PER_DAY = 24 * 60 * 60 * 1000;

  /** Parse a date from ISO (YYYY-MM-DD) or M/D/YYYY; returns null on failure. */
  function parseDate(value) {
    if (!value) return null;
    const s = String(value).trim();
    if (!s) return null;
    // ISO first (unambiguous, preferred).
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      const d = new Date(+m[1], +m[2] - 1, +m[3]);
      return isNaN(d.getTime()) ? null : d;
    }
    // Then M/D/YYYY or M-D-YYYY (US-style, the common Cameo/Excel default).
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      let year = +m[3];
      if (year < 100) year += 2000;
      const d = new Date(year, +m[1] - 1, +m[2]);
      return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  /** Snap a date to the Monday that begins its week (local time). */
  function startOfWeek(date) {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dow = (d.getDay() + 6) % 7; // 0 = Monday
    d.setDate(d.getDate() - dow);
    return d;
  }

  /** Whole-week offset of `date` from the project start (can be negative). */
  function weekIndex(date, projectStart) {
    const a = startOfWeek(projectStart).getTime();
    const b = startOfWeek(date).getTime();
    return Math.round((b - a) / (7 * MS_PER_DAY));
  }

  /** Inverse of weekIndex: the Monday date for a given week index. */
  function weekToDate(index, projectStart) {
    const d = startOfWeek(projectStart);
    d.setDate(d.getDate() + index * 7);
    return d;
  }

  /* ---------------------------------------------------------------------------
   * SECTION 4 — dependency ordering (topological sort)
   * A task may depend on others (precedence). We order tasks so every task
   * comes after its prerequisites. Among tasks that are equally "ready", we
   * break ties by the user's rank (lower rank = higher priority). This is the
   * mechanism behind "bump based on precedence": rank chooses who goes first,
   * but precedence can never be violated.
   * -------------------------------------------------------------------------*/

  /**
   * Kahn's algorithm with a rank-aware tie-break.
   * @returns {{order: string[], hadCycle: boolean}}
   *   `order` always contains every task; if a cycle exists the offending
   *   tasks are appended in rank order so nothing is dropped.
   */
  function topoOrder(tasks) {
    const byId = {};
    const indeg = {};
    const adj = {};
    tasks.forEach(function (t) {
      byId[t.uid] = t;
      indeg[t.uid] = 0;
      adj[t.uid] = [];
    });
    tasks.forEach(function (t) {
      (t.deps || []).forEach(function (d) {
        if (byId[d]) {
          adj[d].push(t.uid); // edge prerequisite -> dependent
          indeg[t.uid]++;
        }
      });
    });

    const cmp = function (a, b) {
      const ra = byId[a].rank;
      const rb = byId[b].rank;
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };

    let ready = tasks
      .filter(function (t) {
        return indeg[t.uid] === 0;
      })
      .map(function (t) {
        return t.uid;
      });

    const order = [];
    while (ready.length) {
      ready.sort(cmp);
      const u = ready.shift();
      order.push(u);
      adj[u].forEach(function (v) {
        if (--indeg[v] === 0) ready.push(v);
      });
    }

    const hadCycle = order.length !== tasks.length;
    if (hadCycle) {
      const seen = {};
      order.forEach(function (u) {
        seen[u] = true;
      });
      tasks
        .slice()
        .sort(function (a, b) {
          return cmp(a.uid, b.uid);
        })
        .forEach(function (t) {
          if (!seen[t.uid]) order.push(t.uid);
        });
    }
    return { order: order, hadCycle: hadCycle };
  }

  /* ---------------------------------------------------------------------------
   * SECTION 5 — the scheduler
   * Two modes:
   *   'forward' — capacity-constrained list scheduling. Tasks are placed as
   *               early as possible in priority order, across `lanes` parallel
   *               work streams (the unknown build-team capacity), always after
   *               their prerequisites finish. This is what makes reprioritizing
   *               visibly reshape the roadmap.
   *   'fixed'   — honour the target completion dates from the CSV; each bar ends
   *               on its completion week and runs backwards by its duration.
   * -------------------------------------------------------------------------*/

  /**
   * Forward, capacity-constrained schedule.
   * @param {Object[]} tasks  {uid, duration, deps[], rank, type, anchorStart?}
   * @param {number}   lanes  Number of parallel work streams (>=1).
   * @returns {Object<string,{startWeek:number,finishWeek:number,lane:number}>}
   */
  function scheduleForward(tasks, lanes) {
    const byId = {};
    tasks.forEach(function (t) {
      byId[t.uid] = t;
    });
    const order = topoOrder(tasks).order;
    const laneFree = new Array(Math.max(1, lanes | 0 || 1)).fill(0);
    const finish = {};
    const result = {};

    order.forEach(function (uid) {
      const t = byId[uid];
      // Earliest start = after every prerequisite finishes, and after any anchor
      // the user has pinned by dragging the bar.
      let earliest = 0;
      (t.deps || []).forEach(function (d) {
        if (finish[d] != null) earliest = Math.max(earliest, finish[d]);
      });
      if (t.anchorStart != null) earliest = Math.max(earliest, t.anchorStart);

      if (t.type === 'milestone') {
        // Milestones are zero-duration markers; they do not consume capacity.
        result[uid] = { startWeek: earliest, finishWeek: earliest, lane: -1 };
        finish[uid] = earliest;
        return;
      }

      const dur = Math.max(0, t.duration || 0);
      // Pack into the latest-freeing lane that is already free by `earliest`
      // (keeps lanes dense); otherwise take the lane that frees soonest and wait.
      let packLane = -1;
      for (let l = 0; l < laneFree.length; l++) {
        if (laneFree[l] <= earliest) {
          if (packLane < 0 || laneFree[l] > laneFree[packLane]) packLane = l;
        }
      }
      let lane = packLane;
      if (lane < 0) {
        lane = 0;
        for (let l = 1; l < laneFree.length; l++) {
          if (laneFree[l] < laneFree[lane]) lane = l;
        }
      }
      const start = Math.max(earliest, laneFree[lane]);
      const fin = start + dur;
      result[uid] = { startWeek: start, finishWeek: fin, lane: lane };
      laneFree[lane] = fin;
      finish[uid] = fin;
    });
    return result;
  }

  /**
   * Fixed-date schedule from target completion weeks.
   * Tasks with no completion date fall back to their forward position so they
   * are never silently stacked at week zero.
   */
  function scheduleFixed(tasks, lanes) {
    const forward = scheduleForward(tasks, lanes);
    const result = {};
    tasks.forEach(function (t) {
      const dur = t.type === 'milestone' ? 0 : Math.max(0, t.duration || 0);
      if (t.completionWeek != null) {
        result[t.uid] = {
          startWeek: t.completionWeek - dur,
          finishWeek: t.completionWeek,
          lane: -1,
        };
      } else {
        result[t.uid] = forward[t.uid];
      }
    });
    return result;
  }

  /** Dispatch to the requested schedule mode. */
  function schedule(tasks, opts) {
    opts = opts || {};
    const mode = opts.mode || 'forward';
    const lanes = opts.lanes || 1;
    return mode === 'fixed'
      ? scheduleFixed(tasks, lanes)
      : scheduleForward(tasks, lanes);
  }

  /* ---------------------------------------------------------------------------
   * SECTION 6 — validation (surfaced as gentle warnings, never blocking)
   * -------------------------------------------------------------------------*/

  /**
   * Produce human-readable warnings: unknown dependency references, cycles, and
   * rank/precedence conflicts (a task ranked above one of its prerequisites).
   * @returns {{level:string, message:string}[]}
   */
  function validate(tasks) {
    const warnings = [];
    const byId = {};
    tasks.forEach(function (t) {
      byId[t.uid] = t;
    });

    tasks.forEach(function (t) {
      (t.deps || []).forEach(function (d) {
        if (!byId[d]) {
          warnings.push({
            level: 'error',
            message:
              'Task "' + t.name + '" depends on unknown id "' + d + '".',
          });
        }
      });
    });

    if (topoOrder(tasks).hadCycle) {
      warnings.push({
        level: 'error',
        message:
          'A circular dependency was detected. Affected tasks are scheduled by rank until the loop is broken.',
      });
    }

    tasks.forEach(function (t) {
      (t.deps || []).forEach(function (d) {
        const pre = byId[d];
        if (pre && pre.rank > t.rank) {
          warnings.push({
            level: 'warn',
            message:
              '"' +
              t.name +
              '" is ranked above its prerequisite "' +
              pre.name +
              '". Precedence still holds, but the order looks inconsistent.',
          });
        }
      });
    });
    return warnings;
  }

  /** New rank order (as a uid->rank map) that respects all precedence links. */
  function autoFixRanks(tasks) {
    const order = topoOrder(tasks).order;
    const map = {};
    order.forEach(function (uid, i) {
      map[uid] = i + 1;
    });
    return map;
  }

  /* ---------------------------------------------------------------------------
   * SECTION 7 — color & contrast
   * Status drives bar color. To keep labels legible on any status color, we pick
   * black or white text from the fill's relative luminance (WCAG-style), so the
   * caller never has to hand-tune readability.
   * -------------------------------------------------------------------------*/

  /** Parse "#RRGGBB" (or "#RGB") into [r,g,b] 0-255; returns null on failure. */
  function hexToRgb(hex) {
    if (!hex) return null;
    var h = String(hex).trim().replace(/^#/, '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    if (!/^[0-9a-fA-F]{6}$/.test(h)) return null;
    return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  }

  /** Relative luminance (0=black,1=white) using the sRGB coefficients. */
  function relLuminance(hex) {
    var rgb = hexToRgb(hex);
    if (!rgb) return 0;
    var a = rgb.map(function (v) {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
    });
    return 0.2126 * a[0] + 0.7152 * a[1] + 0.0722 * a[2];
  }

  /** Choose readable text ('#fff' or a dark ink) for a given background fill. */
  function readableText(hex) {
    return relLuminance(hex) > 0.5 ? '#16202B' : '#FFFFFF';
  }

  /** Format a Date as M/D/YYYY to round-trip the uploaded deliveryDate style. */
  function usDate(d) {
    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
  }

  /** Coerce common truthy spellings ("TRUE","yes","1") to a boolean. */
  function parseBool(v) {
    if (v === true) return true;
    if (v == null) return false;
    return /^(true|yes|y|1|t)$/i.test(String(v).trim());
  }

  /* ---------------------------------------------------------------------------
   * Public surface
   * -------------------------------------------------------------------------*/
  const GanttCore = {
    stripBOM: stripBOM,
    detectDelimiter: detectDelimiter,
    parseCSV: parseCSV,
    parseTable: parseTable,
    toCSV: toCSV,
    parseNumber: parseNumber,
    priorityToRank: priorityToRank,
    parseDate: parseDate,
    startOfWeek: startOfWeek,
    weekIndex: weekIndex,
    weekToDate: weekToDate,
    topoOrder: topoOrder,
    schedule: schedule,
    scheduleForward: scheduleForward,
    scheduleFixed: scheduleFixed,
    validate: validate,
    autoFixRanks: autoFixRanks,
    hexToRgb: hexToRgb,
    relLuminance: relLuminance,
    readableText: readableText,
    usDate: usDate,
    parseBool: parseBool,
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = GanttCore; // Node.js (tests)
  } else {
    root.GanttCore = GanttCore; // Browser (app)
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);
