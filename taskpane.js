Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run").onclick = callGroqAI;
    }
});

async function callGroqAI() {
    const prompt = document.getElementById("prompt").value;
    const status = document.getElementById("status");

    if (!prompt) { alert("Iltimos, vazifa yozing!"); return; }

    status.style.display = "block";
    status.innerText = "AI o'ylamoqda...";

    // Kalitni bo'laklab yashirish
    const k1 = "gsk_E3fp4aq";
    const k2 = "BioKqfmIoObtvW";
    const k3 = "Gdyb3FY6V0O6R3BX";
    const k4 = "MyCSmPEAXDxzONa";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + k1 + k2 + k3 + k4
            },
            body: JSON.stringify({
                model: "mixtral-8x7b-32768",
                messages: [
                    { role: "system", content: "Sen faqat Excel formulasini qaytaradigan mutaxassissan. Ortiqcha gapirma." },
                    { role: "user", content: prompt }
                ]
            })
        });

        const data = await response.json();
        const aiResponse = data.choices[0].message.content.trim();

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            if (aiResponse.startsWith("=")) {
                range.formulas = [[aiResponse]];
            } else {
                range.values = [[aiResponse]];
            }
            await context.sync();
        });

        status.innerText = "Natija: " + aiResponse;

    } catch (e) {
        status.innerText = "Xatolik: " + e.message;
    }
}