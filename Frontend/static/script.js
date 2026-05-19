function fillTopic(topic) {
    document.getElementById("topic").value = topic;
    document.getElementById("topic").focus();
}

async function generatePPT() {

    const topic = document.getElementById("topic").value.trim();
    const slides = document.getElementById("slides").value;
    const theme = document.getElementById("theme").value;

    if (!topic) {
        const statusEl = document.getElementById("status");
        statusEl.textContent = "⚠️ Please enter a topic";
        statusEl.style.color = "#c80000";
        return;
    }

    const button = document.querySelector(".generate-btn");
    const statusEl = document.getElementById("status");

    button.innerHTML = "⏳ Generating...";
    button.disabled = true;
    statusEl.textContent = "Creating your presentation...";
    statusEl.style.color = "#0066cc";

    try {
        const backendUrl = window.location.protocol === 'file:' || window.location.port !== '5000' ? 'http://127.0.0.1:5000' : '';
        const response = await fetch(`${backendUrl}/generate`, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ 
                topic,
                num_slides: parseInt(slides),
                theme
            })
        });

        if (!response.ok) {

            let errorMessage = "Failed to generate PPT";

            try {
                const errData = await response.json();
                errorMessage = errData.error || errorMessage;
            } catch (e) {}

            throw new Error(errorMessage);
        }

        const data = await response.json();
        
        if (!data.success) {
            throw new Error(data.error || "Failed to generate PPT");
        }

        // Trigger download using the URL provided by the backend
        let downloadUrl = data.download_url;
        if (downloadUrl.startsWith('/')) {
            const backendUrl = window.location.protocol === 'file:' || window.location.port !== '5000' ? 'http://127.0.0.1:5000' : '';
            downloadUrl = backendUrl + downloadUrl;
        }
        window.location.href = downloadUrl;

        statusEl.textContent = "✅ PPT Generated Successfully! Downloading...";
        statusEl.style.color = "#007800";

        setTimeout(() => {
            statusEl.textContent = "";
            document.getElementById("topic").value = "";
        }, 3000);

    } catch (error) {

        console.error(error);
        statusEl.textContent = "❌ " + error.message;
        statusEl.style.color = "#c80000";

    } finally {

        button.innerHTML = "✨ Generate Professional PPT";
        button.disabled = false;
    }
}