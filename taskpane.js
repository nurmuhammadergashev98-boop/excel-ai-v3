Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Excel tayyor!");
    }
    document.getElementById("run").onclick = callGroqAI;
});

async function callGroqAI() {
    const promptElement = document.getElementById("prompt");
    const status = document.getElementById("status");

    if (!promptElement.value) {
        status.style.display = "block";
        status.innerText = "Iltimos, vazifani yozing!";
        return;
    }

    status.style.display = "block";
    status.innerText = "AI o'ylamoqda...";

    // Kalitni o'zingizniki bilan tekshiring (gsk_...)
    const apiKey = "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";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + apiKey
            },
            body: JSON.stringify({
                "model": "llama3-8b-8192",
                "messages": [
                    {
                        "role": "system",
                        "content": "Sen faqat Excel formulasini qaytaradigan robotsan. Ortiqcha gapirma."
                    },
                    {
                        "role": "user",
                        "content": promptElement.value
                    }
                ],
                "temperature": 0.5
            })
        });

        const data = await response.json();

        if (!response.ok) {
            console.error("Groq xatosi:", data);
            throw new Error(data.error ? data.error.message : "API xatosi");
        }

        const aiResponse = data.choices[0].message.content.trim();
        status.innerText = "AI javobi: " + aiResponse;

        // Excelga yozish qismi
        if (typeof Excel !== 'undefined') {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                if (aiResponse.startsWith("=")) {
                    range.formulas = [[aiResponse]];
                } else {
                    range.values = [[aiResponse]];
                }
                await context.sync();
            });
        }

    } catch (error) {
        console.error("Xatolik:", error);
        status.innerText = "Xatolik: " + error.message;
    }
}