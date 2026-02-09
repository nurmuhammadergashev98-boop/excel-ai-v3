Office.onReady();

async function callGroqAI() {
    const promptText = document.getElementById("prompt").value;
    const statusBox = document.getElementById("status-container");
    const statusText = document.getElementById("status-text");
    const aiTalk = document.getElementById("ai-talk");

    if (!promptText) return;

    statusBox.style.display = "block";
    statusText.innerText = "Nur AI o'ylamoqda...";
    aiTalk.innerText = "";

    const apiKey = "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";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + apiKey },
            body: JSON.stringify({
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {
                        "role": "system",
                        "content": `Sen Excel boshqaruvchi yordamchisan. Foydalanuvchi buyrug'iga qarab faqat JSON qaytar.
                        Format: {"reply": "AIdan qisqa gap", "action": "write/format/formula", "value": "qiymat", "color": "rang (agar format bo'lsa)"}
                        Misol: "A1 ga salom deb yoz" -> {"reply": "Bajarildi!", "action": "write", "cell": "A1", "value": "salom"}
                        Misol: "A1:A10 ni qizil qil" -> {"reply": "Bo'yab qo'ydim!", "action": "format", "cell": "A1:A10", "color": "red"}`
                    },
                    { "role": "user", "content": promptText }
                ],
                "response_format": { "type": "json_object" }
            })
        });

        const data = await response.json();
        const res = JSON.parse(data.choices[0].message.content);

        aiTalk.innerText = res.reply;

        await Excel.run(async (context) => {
            let range;
            if (res.cell) {
                range = context.workbook.worksheets.getActiveWorksheet().getRange(res.cell);
            } else {
                range = context.workbook.getSelectedRange();
            }

            if (res.action === "write" || res.action === "formula") {
                if (res.value.startsWith("=")) range.formulas = [[res.value]];
                else range.values = [[res.value]];
            } 
            
            if (res.action === "format" && res.color) {
                range.format.fill.color = res.color;
            }

            await context.sync();
        });

        statusText.innerText = "Tayyor!";

    } catch (error) {
        statusText.innerText = "Xatolik yuz berdi.";
        console.error(error);
    }
}

document.getElementById("run").onclick = callGroqAI;