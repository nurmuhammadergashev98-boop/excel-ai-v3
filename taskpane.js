Office.onReady(() => {
    const chat = document.getElementById("chat-container");
    const input = document.getElementById("prompt");
    const btn = document.getElementById("sendBtn");
    const keyInput = document.getElementById("keyInput");

    if (localStorage.getItem("groq_key")) keyInput.value = localStorage.getItem("groq_key");

    function addMsg(t, c, id = null) {
        const d = document.createElement("div");
        if(id) d.id = id;
        d.className = `bubble ${c}`;
        d.innerText = t;
        chat.appendChild(d);
        chat.parentNode.scrollTop = chat.parentNode.scrollHeight;
        return d;
    }

    btn.onclick = async () => {
        const text = input.value.trim();
        const key = keyInput.value.trim();

        if (!text || !key) return;
        localStorage.setItem("groq_key", key);

        addMsg(text, 'user');
        input.value = "";

        // O'ylash animatsiyasini chiqarish
        const loadingMsg = addMsg("Nur AI o'ylamoqda", "ai dots", "loading");
        document.getElementById("status-dot").style.color = "orange";

        try {
            const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
                body: JSON.stringify({
                    "model": "llama-3.3-70b-versatile",
                    "messages": [
                        { "role": "system", "content": `Sen professional Excel yordamchisisan. FAQAT O'ZBEK TILIDA javob ber. 
                        FAQAT JSON QAYTAR: {"reply": "...", "action": "write/clear/format/formula", "cell": "A1", "data": [[]], "color": "#hex"}.
                        - write: ma'lumot yozish.
                        - formula: Excel formulalarini yozish (masalan: "=SUM(A1:A10)").
                        - format: rang berish (color maydonida HEX kod bo'lsin).
                        - clear: varaqni tozalash.` },
                        { "role": "user", "content": text }
                    ],
                    "response_format": { "type": "json_object" }
                })
            });

            const json = await response.json();
            if (!response.ok) throw new Error(json.error.message);

            const res = JSON.parse(json.choices[0].message.content);
            
            // Animatsiyani olib tashlash va haqiqiy javobni yozish
            loadingMsg.remove();
            addMsg(res.reply, "ai");

            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                
                if (res.action === "write" && res.data) {
                    const range = sheet.getRange(res.cell || "A1").getResizedRange(res.data.length - 1, res.data[0].length - 1);
                    range.values = res.data;
                } 
                else if (res.action === "formula" && res.data) {
                    const range = sheet.getRange(res.cell || "A1").getResizedRange(res.data.length - 1, res.data[0].length - 1);
                    range.formulas = res.data;
                }
                else if (res.action === "format") {
                    const range = sheet.getRange(res.cell || "A1");
                    range.format.fill.color = res.color || "#FFFF00"; // Standart sariq
                }
                else if (res.action === "clear") {
                    sheet.getUsedRange().clear();
                }

                await context.sync();
            });

        } catch (e) {
            if(document.getElementById("loading")) document.getElementById("loading").remove();
            addMsg("Kechirasiz, buyruqni tushunmadim yoki xatolik yuz berdi: " + e.message, "ai");
        } finally {
            document.getElementById("status-dot").style.color = "#00ff00";
        }
    };
});