// graphs.js - Handles rendering static charts for finished sessions
// Improvements:
// - Responsive to taskpane width/height (ResizeObserver + maintainAspectRatio:false)
// - Taller canvases for better visibility
// - Interval-based WPM calculation (more accurate)
// - Visual clipping for extreme WPM spikes but tooltips show real values
// - Uses numeric arrays (not strings) for Chart.js data

let activeCharts = {
  progress: null,
  speed: null
};

// Helper: safe destroy chart if exists
function destroyChartIfExists(c) {
  try { if (c && typeof c.destroy === "function") c.destroy(); } catch (e) { /* ignore */ }
}

// Utility: create canvas with consistent sizing
function makeCanvas(heightPx = 280) {
  const canvas = document.createElement("canvas");
  canvas.style.display = "block";
  canvas.style.width = "100%";
  canvas.style.height = `${heightPx}px`; // visible height in CSS pixels
  // ensure the backing store DPI is handled automatically by Chart.js
  return canvas;
}

// Primary renderer called from taskpane.js
window.renderSessionGraphs = function (session) {
  // 1) Find or create container
  let container = document.getElementById("sessionDetailGraphArea");

  if (!container) {
    const summaryDiv = document.getElementById("pastSessionsContainer");
    container = document.createElement("div");
    container.id = "sessionDetailGraphArea";
    container.style.marginTop = "20px";
    container.style.padding = "15px";
    container.style.background = "#fff";
    container.style.borderRadius = "8px";
    container.style.border = "1px solid #e0e0e0";
    // make container flex column so canvases can size nicely
    container.style.display = "flex";
    container.style.flexDirection = "column";
    container.style.gap = "18px";
    if (summaryDiv && summaryDiv.parentNode) summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    else document.body.appendChild(container);
  }

  // clear previous contents
  container.innerHTML = "";
  const title = document.createElement("h4");
  title.textContent = "Session Analysis";
  title.style.margin = "0 0 8px 0";
  container.appendChild(title);

  // 2) Validate
  if (!session || !Array.isArray(session.events) || session.events.length < 2) {
    const note = document.createElement("div");
    note.innerHTML = "<i>Not enough data points to generate a graph.</i>";
    container.appendChild(note);
    return;
  }

  // 3) Process events -> datasets
  // We'll compute:
  // - labels: cumulative time mm:ss from session start
  // - wordsCumulative: estimated words at each event (characters/5)
  // - wpmInterval: words typed between events / minutes between events (gives momentary speed)
  const events = session.events.slice().sort((a, b) => a.t - b.t); // ensure sorted
  const startTime = events[0].t;

  const labels = [];
  const wordsCumulative = [];
  const wpmRaw = []; // actual numeric WPM values for tooltip
  const wpmPlot = []; // possibly clipped values for plotting

  let prevEvent = events[0];
  let prevWords = (prevEvent.c || 0) / 5;

  // Start from index 1 so we can compute interval deltas
  for (let i = 1; i < events.length; i++) {
    const ev = events[i];
    const timeDiffMsFromStart = ev.t - startTime;
    if (timeDiffMsFromStart < 1000) {
      // skip any extremely early events (avoid divide by ~0)
      prevEvent = ev;
      prevWords = (ev.c || 0) / 5;
      continue;
    }

    // label mm:ss
    const totalSeconds = Math.floor(timeDiffMsFromStart / 1000);
    const mm = Math.floor(totalSeconds / 60);
    const ss = totalSeconds % 60;
    labels.push(`${mm}:${ss.toString().padStart(2, "0")}`);

    // cumulative words
    const cumWords = (ev.c || 0) / 5;
    wordsCumulative.push(Number(cumWords.toFixed(1)));

    // interval WPM (words typed between prevEvent and ev) / minutes
    const deltaWords = cumWords - prevWords;
    const deltaMinutes = (ev.t - prevEvent.t) / 60000;
    let intervalWPM = 0;
    if (deltaMinutes > 0) intervalWPM = deltaWords / deltaMinutes;
    // if intervalWPM is negative (possible if chars decreased) clamp to 0
    if (!isFinite(intervalWPM) || intervalWPM < 0) intervalWPM = 0;

    // store raw
    wpmRaw.push(Math.round(intervalWPM));
    prevEvent = ev;
    prevWords = cumWords;
  }

  // If wordsCumulative length differs from wpmRaw length (should be same), ensure alignment
  // (both are created from same loops above; just guard)
  const maxLen = Math.max(labels.length, wordsCumulative.length, wpmRaw.length);
  while (labels.length < maxLen) labels.push("");
  while (wordsCumulative.length < maxLen) wordsCumulative.push(0);
  while (wpmRaw.length < maxLen) wpmRaw.push(0);

  // 4) Determine plotting behaviour for WPM to avoid massive outliers wrecking visualization
  const maxWpm = Math.max(...wpmRaw, 0);
  // Heuristic: if maxWpm is huge (>300), clamp plot to a displayMax while still showing real values in tooltip.
  const DISPLAY_WPM_CAP = 300; // changeable; keeps graph readable
  const displayMax = Math.max( Math.min(DISPLAY_WPM_CAP, Math.ceil(maxWpm/50)*50), 100 ); // at least 100

  for (let i = 0; i < wpmRaw.length; i++) {
    let v = wpmRaw[i];
    // If v > displayMax, plot it at displayMax (so chart stays readable)
    if (v > displayMax) v = displayMax;
    wpmPlot.push(Math.round(v));
  }

  // 5) Create canvases
  // Remove old canvases if present then create fresh ones to avoid sizing issues
  // (destroy charts first)
  destroyChartIfExists(activeCharts.progress);
  destroyChartIfExists(activeCharts.speed);

  // create two canvases with explicit heights so maintainAspectRatio:false works well
  const canvasWords = makeCanvas(300); // taller for better visibility
  const canvasWpm = makeCanvas(300);

  container.appendChild(canvasWords);
  container.appendChild(canvasWpm);

  // 6) Chart config common
  const commonOptions = {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 300 },
    plugins: {
      legend: { display: false },
      tooltip: {
        enabled: true,
        mode: 'index',
        intersect: false,
        // custom tooltip label handled below where needed
      }
    },
    scales: {
      x: {
        ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 12 },
        grid: { display: false }
      }
    }
  };

  // 7) Render Words Over Time (cumulative)
  activeCharts.progress = new Chart(canvasWords.getContext("2d"), {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Words Typed (Cumulative)',
        data: wordsCumulative,
        borderColor: '#0078d7',
        backgroundColor: 'rgba(0, 120, 215, 0.08)',
        fill: true,
        tension: 0.25,
        pointRadius: 3,
        pointHoverRadius: 6,
        borderWidth: 2,
      }]
    },
    options: Object.assign({}, commonOptions, {
      plugins: {
        ...commonOptions.plugins,
        title: { display: true, text: 'Productivity — Words over Time' }
      },
      scales: {
        ...commonOptions.scales,
        y: {
          beginAtZero: true,
          title: { display: true, text: 'Words' },
          ticks: { precision: 0 } // integer-ish ticks
        }
      }
    })
  });

  // 8) Render WPM Over Session (interval WPM)
  // We will show plotted values (possibly clipped) but tooltip will show the raw WPM value.
  activeCharts.speed = new Chart(canvasWpm.getContext("2d"), {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Typing Speed (WPM)',
        data: wpmPlot,
        borderColor: '#d93025',
        backgroundColor: 'rgba(217, 48, 37, 0.08)',
        fill: true,
        tension: 0.25,
        pointRadius: 3,
        pointHoverRadius: 6,
        borderWidth: 2,
      }]
    },
    options: Object.assign({}, commonOptions, {
      plugins: {
        ...commonOptions.plugins,
        title: { display: true, text: 'Speed — WPM over Session' },
        tooltip: {
          ...commonOptions.plugins.tooltip,
          callbacks: {
            label: function(context) {
              const idx = context.dataIndex;
              const plotted = context.dataset.data[idx];
              const raw = (wpmRaw && typeof wpmRaw[idx] !== "undefined") ? wpmRaw[idx] : plotted;
              // If plotted was clipped, show both plotted and true
              if (raw > displayMax) {
                return `WPM: ${raw} (clipped for display)`;
              } else {
                return `WPM: ${raw}`;
              }
            }
          }
        }
      },
      scales: {
        ...commonOptions.scales,
        y: {
          beginAtZero: true,
          title: { display: true, text: 'WPM' },
          suggestedMax: displayMax,
          ticks: {
            callback: function(value) {
              return Number(value).toLocaleString();
            }
          }
        }
      }
    })
  });

  // 9) Resize handling — Chart.js sometimes needs an explicit resize when the containing iframe/pane size changes.
  // We'll attach a ResizeObserver to the container and call chart.resize() when needed.
  if (window._pw_graphs_resize_observer) {
    try { window._pw_graphs_resize_observer.disconnect(); } catch (e) {}
  }
  window._pw_graphs_resize_observer = new ResizeObserver(() => {
    try {
      if (activeCharts.progress) activeCharts.progress.resize();
      if (activeCharts.speed) activeCharts.speed.resize();
    } catch (e) { /* ignore */ }
  });
  try { window._pw_graphs_resize_observer.observe(container); } catch (e) { /* ignore */ }

  // 10) Scroll into view so classroom/teacher sees it after clicking
  try { container.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch (e) {}

  // 11) Return some metadata if caller wants it
  return {
    plottedMaxWPM: Math.max(...wpmPlot, 0),
    rawMaxWPM: maxWpm
  };
};
