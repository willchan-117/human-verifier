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
    let lastChars = session.events[0].c;
    let lastTime = startTime;

    session.events.forEach((event, i) => {
        const timeDiffMs = event.t - startTime;
        if (timeDiffMs < 1000) return;

        const totalSeconds = Math.floor(timeDiffMs / 1000);
        const mm = Math.floor(totalSeconds / 60);
        const ss = totalSeconds % 60;
        labels.push(`${mm}:${ss.toString().padStart(2, "0")}`);

        // --- Correct cumulative words ---
        const cumulativeWords = event.c / 5;
        wordsData.push(cumulativeWords.toFixed(1));

        // --- Correct incremental WPM (prevents 2000+ spikes) ---
        const charsTypedNow = event.c - lastChars;
        const intervalMinutes = (event.t - lastTime) / 60000;

        let wpm = 0;
        if (charsTypedNow > 0 && intervalMinutes > 0) {
            wpm = (charsTypedNow / 5) / intervalMinutes;
        }

        lastChars = event.c;
        lastTime = event.t;

        wpmData.push(Math.round(wpm));
    });

    const canvas1 = document.createElement("canvas");
    canvas1.style.marginBottom = "20px";

    const canvas2 = document.createElement("canvas");

    container.appendChild(canvas1);
    container.appendChild(canvas2);

    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // Words Over Time
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Words Typed (Cumulative)',
                data: wordsData,
                borderColor: '#0078d7',
                backgroundColor: 'rgba(0, 120, 215, 0.1)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Productivity: Words over Time' }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    // WPM Over Time
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Typing Speed (WPM)',
                data: wpmData,
                borderColor: '#d93025',
                backgroundColor: 'rgba(217, 48, 37, 0.1)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Speed: WPM over Session' }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    // 🔥 Fix graph squashing in default Word taskpane
    setTimeout(() => {
        activeCharts.progress.resize();
        activeCharts.speed.resize();
    }, 100);

    container.scrollIntoView({ behavior: 'smooth' });
};
