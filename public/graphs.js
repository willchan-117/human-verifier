// graphs.js - Handles rendering static charts for finished sessions

let activeCharts = {
    progress: null,
    speed: null
};

let lastRenderedSession = null;  // so resize can re-render

// Re-render charts on resize (Word taskpane collapses → canvas becomes tiny)
const observer = new ResizeObserver(() => {
    if (lastRenderedSession) {
        renderSessionGraphs(lastRenderedSession);
    }
});

// This name matches the call in taskpane.js
window.renderSessionGraphs = function (session) {
    lastRenderedSession = session; // save for resize observer

    // 1. Find or create the container
    let container = document.getElementById("sessionDetailGraphArea");

    if (!container) {
        const summaryDiv = document.getElementById("pastSessionsContainer");
        container = document.createElement("div");
        container.id = "sessionDetailGraphArea";
        container.style.marginTop = "20px";
        container.style.background = "#fff";
        container.style.padding = "16px";
        container.style.borderRadius = "10px";
        container.style.border = "1px solid #ddd";
        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);

        observer.observe(container); // observe size changes
    }

    container.innerHTML = "<h4 style='margin-bottom:12px;'>Session Analysis</h4>";

    // 2. Validate data
    if (!session.events || session.events.length < 3) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    // 3. Process Data (correct WPM)
    const labels = [];
    const wordsData = [];
    const wpmData = [];

    const startTime = session.events[0].t;

    let lastTime = startTime;
    let lastChars = session.events[0].c;

    session.events.forEach((event) => {
        const dt = event.t - lastTime;       // interval time
        const dc = event.c - lastChars;      // interval chars

        if (dt <= 50) return; // skip micro-events

        const totalSec = Math.floor((event.t - startTime) / 1000);
        const mm = Math.floor(totalSec / 60);
        const ss = (totalSec % 60).toString().padStart(2, '0');

        labels.push(`${mm}:${ss}`);

        // Words typed in this interval
        const intervalWords = dc / 5;
        const intervalMinutes = dt / 60000;

        // Correct WPM calculation
        const wpm = intervalMinutes > 0 ? intervalWords / intervalMinutes : 0;

        // Cumulative words
        const totalWords = (event.c / 5);

        wordsData.push(totalWords.toFixed(1));
        wpmData.push(Math.max(0, wpm).toFixed(0));

        lastTime = event.t;
        lastChars = event.c;
    });

    // 4. Create canvases
    const canvas1 = document.createElement("canvas");
    canvas1.style.marginBottom = "20px";
    canvas1.style.height = "240px";   // **taller**
    canvas1.style.width = "100%";
    container.appendChild(canvas1);

    const canvas2 = document.createElement("canvas");
    canvas2.style.height = "240px";   // **taller**
    canvas2.style.width = "100%";
    container.appendChild(canvas2);

    // 5. Destroy old charts
    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // 6. Render cumulative words
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: 'line',
        data: {
            labels,
            datasets: [{
                label: "Words Typed (Cumulative)",
                data: wordsData,
                borderColor: "#0078d7",
                backgroundColor: "rgba(0,120,215,0.15)",
                fill: true,
                tension: 0.25
            }]
        },
        options: {
            maintainAspectRatio: false,   // **important**
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: "Words" }
                }
            }
        }
    });

    // 7. Render interval WPM
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Typing Speed (WPM)",
                data: wpmData,
                borderColor: "#d93025",
                backgroundColor: "rgba(217,48,37,0.15)",
                fill: true,
                tension: 0.25
            }]
        },
        options: {
            maintainAspectRatio: false,
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: "WPM" }
                }
            }
        }
    });

    container.scrollIntoView({ behavior: "smooth" });
};
