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
        container.style.padding = "15px";
        container.style.background = "#fff";
        container.style.borderRadius = "8px";
        container.style.border = "1px solid #e0e0e0";
        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    }

    container.innerHTML = "<h4>Session Analysis</h4>";

    if (!session.events || session.events.length < 2) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    const labels = [];
    const wordsData = [];
    const wpmData = [];

    const startTime = session.events[0].t;

    session.events.forEach((event) => {
        const timeDiffMs = event.t - startTime;
        const minutes = timeDiffMs / 60000;

        if (timeDiffMs < 1000) return;

        const totalSeconds = Math.floor(timeDiffMs / 1000);
        const mm = Math.floor(totalSeconds / 60);
        const ss = totalSeconds % 60;
        labels.push(`${mm}:${ss.toString().padStart(2, '0')}`);

        const estimatedWords = (event.c / 5);
        wordsData.push(estimatedWords);

        const wpm = minutes > 0 ? (estimatedWords / minutes) : 0;
        wpmData.push(Math.round(wpm));
    });

    // === Create canvases with fixed height so they look normal inside narrow taskpane ===
    const canvas1 = document.createElement("canvas");
    canvas1.style.width = "100%";
    canvas1.style.height = "260px";
    canvas1.style.minHeight = "260px";
    canvas1.style.marginBottom = "25px";
    container.appendChild(canvas1);

    const canvas2 = document.createElement("canvas");
    canvas2.style.width = "100%";
    canvas2.style.height = "260px";
    canvas2.style.minHeight = "260px";
    container.appendChild(canvas2);

    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // ===== Words Over Time =====
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Words Typed (Cumulative)",
                data: wordsData,
                borderColor: "#0078d7",
                backgroundColor: "rgba(0,120,215,0.15)",
                tension: 0.35,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,   // IMPORTANT FIX
            plugins: {
                title: { display: true, text: "Productivity: Words Over Time" }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: "Words" }
                }
            }
        }
    });

    // ===== WPM Over Time =====
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Typing Speed (WPM)",
                data: wpmData,
                borderColor: "#d93025",
                backgroundColor: "rgba(217,48,37,0.15)",
                tension: 0.35,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,  // IMPORTANT FIX
            plugins: {
                title: { display: true, text: "Speed: WPM Over Session" }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: "WPM" }
                }
            }
        }
    });

    // FORCE RESIZE — otherwise Word taskpane won't redraw properly
    setTimeout(() => {
        activeCharts.progress.resize();
        activeCharts.speed.resize();
    }, 10);

    container.scrollIntoView({ behavior: "smooth" });
};
