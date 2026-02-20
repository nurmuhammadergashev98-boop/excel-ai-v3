Office.onReady(() => {
    const chat = document.getElementById("chat-container");
    const input = document.getElementById("prompt");
    const btn = document.getElementById("sendBtn");
    const keyInput = document.getElementById("keyInput");

    if (localStorage.getItem("groq_key")) keyInput.value = localStorage.getItem("groq_key");

    // Xabar qo'shish funksiyasi
    function addMsg(t, c, id = null) {
        const d = document.createElement("div");
        if(id) d.id = id;
        d.className = `bubble ${c}`;
        d.innerText = t;
        chat.appendChild(d);
        // Scrollni har doim pastga tushirish
        const container = document.querySelector(".main-container");
        container.scrollTop = container.scrollHeight;
        return d;
    }

    // ASOSIY RUN FUNKSIYASI
    async function runAI() {
        const text = input.value.trim();
        const key = keyInput.value.trim();

        if (!text || !key) return;

        localStorage.setItem("groq_key", key);
        addMsg(text, 'user');
        input.value = ""; // Inputni darhol tozalash

        const loadingMsg = addMsg("Nur AI o'ylamoqda", "ai dots", "loading");
        document.getElementById("status-dot").style.color = "orange";

        try {
            const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
                body: JSON.stringify({
                    "model": "llama-3.3-70b-versatile",
                    "messages": [
                        { "role": "system", "content": `Siz professional Excel mutaxassisiz. FAQAT O'ZBEK TILIDA javob bering.
                        Har doim JSON qaytaring: {"reply": "...", "action": "write/clear/format/formula", "cell": "A1", "data": [[]], "color": "#hex"}.
                        - format: katakni ranglash.
                        - formula: Excel formulasini yozish.
                        - clear: barcha ma'lumotni o'chirish.` },
                        { "role": "user", "content": text }
                    ],
                    "response_format": { "type": "json_object" }
                })
            });

            const json = await response.json();
            const res = JSON.parse(json.choices[0].message.content);
            
            loadingMsg.remove();
            addMsg(res.reply, "ai");

            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                
                if (res.action === "write" && res.data) {
                    const range = sheet.getRange(res.cell || "A1").getResizedRange(res.data.length - 1, res.data[0].length - 1);
                    range.values = res.data;
                    range.format.autofitColumns();
                } 
                else if (res.action === "formula" && res.data) {
                    const range = sheet.getRange(res.cell || "A1").getResizedRange(res.data.length - 1, res.data[0].length - 1);
                    range.formulas = res.data;
                }
                else if (res.action === "format") {
                    sheet.getRange(res.cell || "A1").format.fill.color = res.color || "#FFFF00";
                }
                else if (res.action === "clear") {
                    sheet.getUsedRange().clear();
                }
                await context.sync();
            });

        } catch (e) {
            if(document.getElementById("loading")) document.getElementById("loading").remove();
            addMsg("Xatolik yuz berdi. Iltimos, qaytadan urinib ko'ring.", "ai");
            console.error(e);
        } finally {
            document.getElementById("status-dot").style.color = "#00ff00";
        }
    }

    // ENTER TUGMASINI BOSGANDA YUBORISH
    input.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault(); // Yangi qatorga o'tishni to'xtatadi
            runAI();
        }
    });

    // SAMOLYOTCHANI BOSGANDA YUBORISH
    btn.onclick = runAI;
});