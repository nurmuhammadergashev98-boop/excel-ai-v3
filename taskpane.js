// Taskpane.js - Professional AI Assistant
Office.onReady(() => {
    const chat = document.getElementById("chat-container");
    const input = document.getElementById("prompt");
    const btn = document.getElementById("sendBtn");
    const keyInput = document.getElementById("keyInput");

    if (localStorage.getItem("groq_key")) keyInput.value = localStorage.getItem("groq_key");

    btn.onclick = async () => {
        const text = input.value.trim();
        const key = keyInput.value.trim();
        if (!text || !key) return;

        localStorage.setItem("groq_key", key);
        addMsg(text, 'user');
        input.value = "";

        try {
            const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
                body: JSON.stringify({
                    "model": "llama-3.3-70b-versatile",
                    "messages": [
                        { "role": "system", "content": "Siz professional Excel yordamchisiz. FAQAT O'ZBEK TILIDA javob bering. Har doim JSON qaytaring: {\"reply\": \"javob\", \"action\": \"write/clear/format\", \"cell\": \"A1\", \"data\": [[]], \"color\": \"#hex\"}. 'clear' buyrug'ida butun varaqni tozalang." },
                        { "role": "user", "content": text }
                    ],
                    "response_format": { "type": "json_object" }
                })
            });

            const json = await response.json();
            const res = JSON.parse(json.choices[0].message.content);
            addMsg(res.reply, 'ai');

            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                
                if (res.action === "write" && res.data) {
                    const range = sheet.getRange(res.cell || "A1").getResizedRange(res.data.length - 1, res.data[0].length - 1);
                    range.values = res.data;
                    range.format.autofitColumns();
                } else if (res.action === "clear") {
                    sheet.getUsedRange().clear(); // Haqiqiy o'chirish buyrug'i
                } else if (res.action === "format") {
                    sheet.getRange(res.cell || "A1").format.fill.color = res.color || "yellow";
                }
                await context.sync();
            });
        } catch (e) {
            addMsg("Xato yuz berdi: " + e.message, 'ai');
        }
    };

    function addMsg(t, c) {
        const d = document.createElement("div");
        d.className = `bubble ${c}`;
        d.innerText = t;
        chat.appendChild(d);
        chat.parentNode.scrollTop = chat.parentNode.scrollHeight;
    }
});