// graphs.js - Handles rendering static charts for finished sessions

let activeCharts = {
    progress: null,
    speed: null
};

// This matches safeRenderSessionGraphs()
window.renderSessionGraphs = function (session) {
    let container = document.getElementById("sessionDetailGraphArea");

    if (!container) {
        const summaryDiv = document.getElementById("pastSessionsContainer");
        container = document.createElement("div");
        container.id = "sessionDetailGraphArea";
        container.style.marginTop = "20px";
        container.style.padding = "15px";
        container.style.background = "#fff";
        container.style.borderRadius = "12px";
        container.style.border = "1px solid #e0e0e0";
        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    }

    container.innerHTML = "<h4 style='margin-bottom:12px;'>Session Analysis</h4>";

    if (!session.events || session.events.length < 2) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    // ---- PREPROCESSING DATA ----
    const labels = [];
    const wordsData = [];
    const wpmData = [];

    let startTime = session.events[0].t;
    let lastTime = startTime;
    let lastChars = session.events[0].c;

    session.events.forEach(event => {
        const dt = event.t - startTime;
        if (dt < 1000) return; // skip first split second

        const totalSec = Math.floor(dt / 1000);
        const mm = Math.floor(totalSec / 60);
        const ss = totalSec % 60;
        labels.push(`${mm}:${ss.toString().padStart(2, '0')}`);

        // Words (cumulative)
        const words = event.c / 5;
        wordsData.push(words);

        // INSTANT WPM FIX ✔
        const deltaChars = event.c - lastChars;
        const deltaTimeMin = (event.t - lastTime) / 60000;

        let instantWPM = 0;
        if (deltaTimeMin > 0 && deltaChars > 0) {
            instantWPM = (deltaChars / 5) / deltaTimeMin;
        }

        wpmData.push(Math.round(instantWPM));

        lastChars = event.c;
        lastTime = event.t;
    });

    // ---- CANVASES ----
    const canvas1 = document.createElement("canvas");
    canvas1.height = 240; // ⭐ taller graph
    canvas1.style.marginBottom = "24px";
    container.appendChild(canvas1);

    const canvas2 = document.createElement("canvas");
    canvas2.height = 240; // ⭐ taller graph
    container.appendChild(canvas2);

    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // ---- CHART DEFAULTS ----
    const baseOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                labels: {
                    font: { size: 14 }
                }
            },
            tooltip: {
                backgroundColor: "rgba(0,0,0,0.7)",
                padding: 10,
                bodyFont: { size: 14 }
            }
        },
        scales: {
            x: {
                ticks: { font: { size: 12 } }
            },
            y: {
                beginAtZero: true,
                ticks: { font: { size: 12 } }
            }
        }
    };

    // ---- Words Over Time ----
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Words Typed (Cumulative)",
                data: wordsData,
                borderWidth: 3,
                borderColor: "#1a73e8",
                backgroundColor: "rgba(26,115,232,0.12)",
                fill: true,
                tension: 0.35
            }]
        },
        options: {
            ...baseOptions,
            plugins: {
                ...baseOptions.plugins,
                title: {
                    display: true,
                    text: "Productivity: Words Over Time",
                    font: { size: 18 }
                }
            }
        }
    });

    // ---- WPM Over Time ----
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Typing Speed (Instant WPM)",
                data: wpmData,
                borderWidth: 3,
                borderColor: "#d93025",
                backgroundColor: "rgba(217,48,37,0.12)",
                fill: true,
                tension: 0.35
            }]
        },
        options: {
            ...baseOptions,
            plugins: {
                ...baseOptions.plugins,
                title: {
                    display: true,
                    text: "Speed: Instant WPM",
                    font: { size: 18 }
                }
            },
            scales: {
                ...baseOptions.scales,
                y: {
                    beginAtZero: true,
                    suggestedMax: 120, // ⭐ better for typical typing
                    title: { display: true, text: "WPM" }
                }
            }
        }
    });

    container.scrollIntoView({ behavior: "smooth" });
};
