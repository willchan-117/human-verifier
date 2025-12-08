// graphs.js - Handles rendering static charts for finished sessions

let activeCharts = {
    progress: null,
    speed: null
};

// Resize observer to auto-resize charts when taskpane size changes
let resizeObserver = null;

// This name matches the call in taskpane.js "safeRenderSessionGraphs"
window.renderSessionGraphs = function (session) {

    // ---------------------------------------------------------
    // 1. Create / find container
    // ---------------------------------------------------------
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
        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    }

    container.innerHTML = "<h4>Session Analysis</h4>";

    // ---------------------------------------------------------
    // 2. Validate Data
    // ---------------------------------------------------------
    if (!session.events || session.events.length < 2) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    // ---------------------------------------------------------
    // 3. Prepare Data
    // ---------------------------------------------------------
    const labels = [];
    const wordsData = [];
    const wpmData = [];

    const startTime = session.events[0].t;
    let lastEvent = null;

    session.events.forEach((event) => {
        const t = event.t - startTime;

        if (t < 1000) return;

        const seconds = Math.floor(t / 1000);
        const mm = Math.floor(seconds / 60);
        const ss = seconds % 60;

        labels.push(`${mm}:${ss.toString().padStart(2, '0')}`);

        // Words estimate (5 chars = 1 word)
        const words = event.c / 5;
        wordsData.push(words.toFixed(1));

        // WPM fix: use difference between events (not cumulative)
        if (lastEvent) {
            const diffChars = event.c - lastEvent.c;
            const diffWords = diffChars / 5;

            const diffTimeMinutes = (event.t - lastEvent.t) / 60000;

            let wpm = diffTimeMinutes > 0 ? diffWords / diffTimeMinutes : 0;

            // Clean out spikes caused by bursts
            if (wpm > 300) wpm = 300;

            wpmData.push(Math.round(wpm));
        } else {
            wpmData.push(0);
        }

        lastEvent = event;
    });

    // ---------------------------------------------------------
    // 4. Create canvases dynamically
    // ---------------------------------------------------------
    const canvas1 = document.createElement("canvas");
    const canvas2 = document.createElement("canvas");

    // Make charts tall & responsive
    canvas1.style.height = "260px";
    canvas2.style.height = "260px";

    container.appendChild(canvas1);
    container.appendChild(canvas2);

    // ---------------------------------------------------------
    // 5. Destroy old charts
    // ---------------------------------------------------------
    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // ---------------------------------------------------------
    // 6. Build Chart.js Config
    // ---------------------------------------------------------
    const baseOptions = {
        responsive: true,
        maintainAspectRatio: false,
        tension: 0.35,
        borderWidth: 2,
        plugins: {
            legend: {
                display: false
            }
        },
        scales: {
            x: {
                ticks: { maxRotation: 0, autoSkip: true }
            },
            y: {
                beginAtZero: true
            }
        }
    };

    // ---------------- Words Over Time ----------------
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Words Typed",
                data: wordsData,
                borderColor: "#0078d7",
                backgroundColor: "rgba(0, 120, 215, 0.12)",
                fill: true
            }]
        },
        options: {
            ...baseOptions,
            plugins: {
                title: { display: true, text: "Words Over Time" }
            }
        }
    });

    // ---------------- WPM Over Time ----------------
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "WPM",
                data: wpmData,
                borderColor: "#d93025",
                backgroundColor: "rgba(217, 48, 37, 0.12)",
                fill: true
            }]
        },
        options: {
            ...baseOptions,
            plugins: {
                title: { display: true, text: "Typing Speed (WPM)" }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    suggestedMax: 120
                }
            }
        }
    });

    // ---------------------------------------------------------
    // 7. Auto-resize on taskpane width change
    // ---------------------------------------------------------
    if (resizeObserver) resizeObserver.disconnect();

    resizeObserver = new ResizeObserver(() => {
        if (activeCharts.progress) activeCharts.progress.resize();
        if (activeCharts.speed) activeCharts.speed.resize();
    });

    resizeObserver.observe(container);

    // ---------------------------------------------------------
    // 8. Scroll into view
    // ---------------------------------------------------------
    container.scrollIntoView({ behavior: "smooth" });
};
