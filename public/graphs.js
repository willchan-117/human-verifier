// graphs.js - Handles rendering static charts for finished sessions

let activeCharts = {
    progress: null,
    speed: null
};

window.renderSessionGraphs = function (session) {
    let container = document.getElementById("sessionDetailGraphArea");

    if (!container) {
        const summaryDiv = document.getElementById("pastSessionsContainer");
        container = document.createElement("div");
        container.id = "sessionDetailGraphArea";
        container.style.marginTop = "20px";
        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    }

    container.innerHTML = `
        <h4 style="margin-bottom:12px;">Session Analysis</h4>
        <div id="wordsChartContainer" style="height:260px; margin-bottom:25px;">
            <canvas id="wordsChart"></canvas>
        </div>
        <div id="wpmChartContainer" style="height:260px;">
            <canvas id="wpmChart"></canvas>
        </div>
    `;

    if (!session.events || session.events.length < 2) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    // Process Data
    const labels = [];
    const wordsData = [];
    const wpmData = [];

    const events = session.events;
    const startTime = events[0].t;

    for (let i = 1; i < events.length; i++) {
        const prev = events[i - 1];
        const curr = events[i];

        const deltaChars = curr.c - prev.c;
        const deltaMs = curr.t - prev.t;

        if (deltaMs <= 200) continue; // skip tiny intervals

        const minutesSinceStart = (curr.t - startTime) / 60000;
        const totalSeconds = Math.floor((curr.t - startTime) / 1000);

        const mm = Math.floor(totalSeconds / 60);
        const ss = totalSeconds % 60;

        labels.push(`${mm}:${ss.toString().padStart(2, "0")}`);

        // Cumulative words (correct)
        const words = curr.c / 5;
        wordsData.push(words.toFixed(1));

        // FIXED WPM (interval-based)
        const minutes = deltaMs / 60000;
        const intervalWords = deltaChars / 5;

        const wpm = minutes > 0 ? intervalWords / minutes : 0;
        wpmData.push(wpm.toFixed(0));
    }

    // Destroy old charts
    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    const wordsCtx = document.getElementById("wordsChart").getContext("2d");
    const wpmCtx = document.getElementById("wpmChart").getContext("2d");

    // Create charts
    activeCharts.progress = new Chart(wordsCtx, {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Words Typed (Cumulative)",
                data: wordsData,
                borderColor: "#0078d7",
                backgroundColor: "rgba(0, 120, 215, 0.15)",
                fill: true,
                tension: 0.35,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    activeCharts.speed = new Chart(wpmCtx, {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Typing Speed (WPM)",
                data: wpmData,
                borderColor: "#d93025",
                backgroundColor: "rgba(217, 48, 37, 0.15)",
                fill: true,
                tension: 0.35,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    // Force charts to resize when the taskpane changes width
    new ResizeObserver(() => {
        if (activeCharts.progress) activeCharts.progress.resize();
        if (activeCharts.speed) activeCharts.speed.resize();
    }).observe(container);
};
