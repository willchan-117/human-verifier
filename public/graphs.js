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
        container.style.borderRadius = "10px";
        container.style.border = "1px solid #ddd";
        container.style.boxShadow = "0 2px 6px rgba(0,0,0,0.07)";
        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    }

    container.innerHTML = "<h4 style='margin-bottom:10px;'>Session Analysis</h4>";

    if (!session.events || session.events.length < 2) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    // ----- Process Data -----
    const labels = [];
    const wordsData = [];
    const wpmData = [];

    const startTime = session.events[0].t;
    let lastWords = 0;
    let lastTime = startTime;

    session.events.forEach(event => {
        const timeDiffMs = event.t - startTime;
        const totalSeconds = Math.floor(timeDiffMs / 1000);
        const mm = Math.floor(totalSeconds / 60);
        const ss = String(totalSeconds % 60).padStart(2, "0");

        const words = event.c / 5;
        wordsData.push(words);

        labels.push(`${mm}:${ss}`);

        const dtMinutes = (event.t - lastTime) / 60000;
        const wordDelta = words - lastWords;

        let wpm = 0;
        if (dtMinutes > 0.01) {
            wpm = (wordDelta / dtMinutes);
        }

        wpmData.push(Math.round(wpm));

        lastWords = words;
        lastTime = event.t;
    });

    // ----- Canvas -----
    const canvas1 = document.createElement("canvas");
    const canvas2 = document.createElement("canvas");

    // Force visible height
    canvas1.height = 300;
    canvas2.height = 300;

    container.appendChild(canvas1);
    container.appendChild(canvas2);

    // Destroy old charts
    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // ----- Words Chart -----
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: "line",
        data: {
            labels: labels,
            datasets: [{
                label: "Words Typed (Cumulative)",
                data: wordsData,
                borderColor: "#1B84FF",
                backgroundColor: "rgba(27, 132, 255, 0.15)",
                fill: true,
                tension: 0.25,
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    // ----- WPM Chart -----
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: "line",
        data: {
            labels: labels,
            datasets: [{
                label: "Typing Speed (WPM)",
                data: wpmData,
                borderColor: "#E63946",
                backgroundColor: "rgba(230, 57, 70, 0.15)",
                fill: true,
                tension: 0.25,
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    container.scrollIntoView({ behavior: "smooth" });
};
