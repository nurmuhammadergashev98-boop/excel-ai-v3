Office.onReady(() => {
    console.log("Office Ready!");
    
    // Elementlarni olish
    const chat = document.getElementById("chat-container");
    const input = document.getElementById("prompt");
    const btn = document.getElementById("sendBtn");
    const counter = document.getElementById("counter");
    const keyInput = document.getElementById("keyInput"); 
    let actionsDone = 0;

    // Kalitni eslab qolish
    if (localStorage.getItem("groq_key")) {
        keyInput.value = localStorage.getItem("groq_key");
    }

    // Xabar qo'shish funksiyasi
    function addMsg(t, c) {
        const d = document.createElement("div");
        d.className = `bubble ${c}`;
        d.innerText = t;
        chat.appendChild(d);
        // Scrollni pastga tushirish
        const scrollContainer = chat.closest('.dashboard-content') || chat.parentNode;
        scrollContainer.scrollTop = scrollContainer.scrollHeight;
    }

    // Asosiy ishga tushirish funksiyasi
    async function run() {
        const text = input.value.trim();
        const key = keyInput.value.trim();

        if (!text) return;
        if (!key) {
            addMsg("Iltimos, avval API kalitini kiriting!", 'ai');
            return;
        }

        localStorage.setItem("groq_key", key);
        addMsg(text, 'user');
        input.value = "";
        document.getElementById("status-dot").style.color = "orange";

        try {
            const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
                method: "POST",
                headers: { 
                    "Content-Type": "application/json", 
                    "Authorization": "Bearer " + key 
                },
                body: JSON.stringify({
                    "model": "llama-3.3-70b-versatile",
                    "messages": [
                        { "role": "system", "content": "Sen professional Excel yordamchisisan. Faqat JSON qaytar: {\"reply\": \"...\", \"action\": \"write\", \"cell\": \"A1\", \"data\": [[]]}" },
                        { "role": "user", "content": text }
                    ],
                    "response_format": { "type": "json_object" }
                })
            });

            if (!response.ok) throw new Error("API ulanishda xato: " + response.status);

            const json = await response.json();
            const res = JSON.parse(json.choices[0].message.content);
            
            addMsg(res.reply, 'ai');

            if (res.action === "write" && res.data) {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const startRange = sheet.getRange(res.cell || "A1");
                    const targetRange = startRange.getResizedRange(res.data.length - 1, res.data[0].length - 1);
                    targetRange.values = res.data;
                    await context.sync();
                    actionsDone++;
                    counter.innerText = actionsDone;
                });
            }
        } catch (e) {
            addMsg("Xato: " + e.message, 'ai');
            console.error("Run Error:", e);
        } finally {
            document.getElementById("status-dot").style.color = "#00ff00";
        }
    }

    // Tugmalarga hodisalarni biriktirish
    if (btn) {
        btn.onclick = (e) => {
            e.preventDefault();
            run();
        };
    }

    if (input) {
        input.onkeydown = (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                run();
            }
        };
    }
});