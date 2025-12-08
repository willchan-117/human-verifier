// graphs.js - Handles rendering static charts for finished sessions

let activeCharts = {
    progress: null,
    speed: null
};

// This name matches the call in taskpane.js "safeRenderSessionGraphs"
window.renderSessionGraphs = function (session) {

    // 1. Find or create the container.
    let container = document.getElementById("sessionDetailGraphArea");
    
    if (!container) {
        const summaryDiv = document.getElementById("pastSessionsContainer");
        container = document.createElement("div");
        container.id = "sessionDetailGraphArea";

        // KEY FIX: Force reliable layout height
        container.style.marginTop = "20px";
        container.style.padding = "15px";
        container.style.background = "#fff";
        container.style.borderRadius = "8px";
        container.style.border = "1px solid #e0e0e0";
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "25px";

        summaryDiv.parentNode.insertBefore(container, summaryDiv.nextSibling);
    }

    container.innerHTML = "<h4>Session Analysis</h4>";

    // 2. Validate data
    if (!session.events || session.events.length < 2) {
        container.innerHTML += "<i>Not enough data points to generate a graph.</i>";
        return;
    }

    // 3. Process data
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
        labels.push(`${mm}:${ss.toString().padStart(2, "0")}`);

        const estimatedWords = event.c / 5;
        wordsData.push(estimatedWords.toFixed(1));

        const wpm = minutes > 0 ? (estimatedWords / minutes) : 0;
        wpmData.push(wpm.toFixed(0));
    });

    // 4. Create canvases with forced height
    const canvas1 = document.createElement("canvas");
    canvas1.style.height = "220px";        // ⭐ FIX
    canvas1.style.minHeight = "220px";     // ⭐ FIX
    container.appendChild(canvas1);

    const canvas2 = document.createElement("canvas");
    canvas2.style.height = "220px";        // ⭐ FIX
    canvas2.style.minHeight = "220px";     // ⭐ FIX
    container.appendChild(canvas2);

    // 5. Destroy old charts
    if (activeCharts.progress) activeCharts.progress.destroy();
    if (activeCharts.speed) activeCharts.speed.destroy();

    // Common Chart.js options
    const baseOptions = {
        responsive: true,
        maintainAspectRatio: false,   // ⭐ MOST IMPORTANT FIX
        animation: false,
        scales: {
            y: { beginAtZero: true }
        },
        plugins: {
            legend: { display: false }
        }
    };

    // 6. Words graph
    activeCharts.progress = new Chart(canvas1.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Words Typed",
                data: wordsData,
                borderColor: "#0078d7",
                backgroundColor: "rgba(0, 120, 215, 0.1)",
                tension: 0.25,
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

    // 7. WPM graph
    activeCharts.speed = new Chart(canvas2.getContext("2d"), {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "WPM",
                data: wpmData,
                borderColor: "#d93025",
                backgroundColor: "rgba(217, 48, 37, 0.1)",
                tension: 0.25,
                fill: true
            }]
        },
        options: {
            ...baseOptions,
            plugins: {
                title: { display: true, text: "WPM Over Time" }
            }
        }
    });

    container.scrollIntoView({ behavior: "smooth" });
};
