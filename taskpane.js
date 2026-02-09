Office.onReady();

const chatArea = document.getElementById("chat-area");
const promptInput = document.getElementById("prompt");
const statusIndicator = document.getElementById("status-indicator");

async function callGroqAI() {
    const text = promptInput.value;
    if (!text) return;

    // Foydalanuvchi xabarini ekranga chiqarish
    addMessage(text, 'user-msg');
    promptInput.value = "";
    statusIndicator.innerText = "typing...";

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
                        "content": `Sen Excel boshqaruvchi yordamchisan. Faqat JSON qaytar.
                        Buyruq turlari:
                        1. write: Katakka matn/son yozish
                        2. format: Rang berish (fontColor, fillColor)
                        3. formula: Excel formulasi qo'yish
                        Format: {"reply": "gap", "action": "write/format/formula", "cell": "A1", "value": "...", "color": "red/#FF0000"}`
                    },
                    { "role": "user", "content": text }
                ],
                "response_format": { "type": "json_object" }
            })
        });

        const data = await response.json();
        const res = JSON.parse(data.choices[0].message.content);

        // AI javobini ekranga chiqarish
        addMessage(res.reply, 'ai-msg');

        // Excel amallarini bajarish
        await Excel.run(async (context) => {
            let range;
            if (res.cell) {
                range = context.workbook.worksheets.getActiveWorksheet().getRange(res.cell);
            } else {
                range = context.workbook.getSelectedRange();
            }

            if (res.action === "write") range.values = [[res.value]];
            if (res.action === "formula") range.formulas = [[res.value]];
            if (res.action === "format") {
                if (res.color) range.format.fill.color = res.color;
            }

            await context.sync();
        });

    } catch (error) {
        addMessage("Xatolik: Serverga ulanib bo'lmadi.", 'ai-msg');
        console.error(error);
    } finally {
        statusIndicator.innerText = "online";
    }
}

function addMessage(text, className) {
    const msgDiv = document.createElement("div");
    msgDiv.className = `message ${className}`;
    msgDiv.innerText = text;
    chatArea.appendChild(msgDiv);
    chatArea.scrollTop = chatArea.scrollHeight; // Avtomatik scroll
}

document.getElementById("run").onclick = callGroqAI;