Office.onReady();

const chat = document.getElementById("chat-container");
const input = document.getElementById("prompt");
const btn = document.getElementById("sendBtn");
const counter = document.getElementById("counter");
const keyInput = document.getElementById("keyInput"); // HTML-dagi yangi maydon
let actionsDone = 0;

// Kalitni brauzer xotirasidan olish (har safar yozmaslik uchun)
if (localStorage.getItem("groq_key")) {
    keyInput.value = localStorage.getItem("groq_key");
}

input.addEventListener("keydown", (e) => { 
    if (e.key === "Enter" && !e.shiftKey) { 
        e.preventDefault(); 
        run(); 
    } 
});

btn.onclick = run;

async function run() {
    const text = input.value.trim();
    const key = keyInput.value.trim();

    if (!text) return;
    if (!key) {
        addMsg("Iltimos, avval API kalitini kiriting!", 'ai');
        return;
    }

    // Kalitni saqlab qo'yish
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
                "messages": [{
                    "role": "system",
                    "content": `Sen professional Excel avtomatizatorisan. Faqat JSON qaytar.
                    Struktura: {"reply": "javob", "action": "write/formula/format", "cell": "A1", "data": [[...]], "color": "#hex"}`
                }, { "role": "user", "content": text }],
                "response_format": { "type": "json_object" }
            })
        });

        const json = await response.json();
        
        if (!response.ok) {
            throw new Error(json.error ? json.error.message : "API xatosi");
        }

        const res = JSON.parse(json.choices[0].message.content);
        addMsg(res.reply, 'ai');

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(res.cell || "A1");

            if (res.action === "write") range.values = res.data;
            if (res.action === "formula") range.formulas = res.data;
            if (res.action === "format") range.format.fill.color = res.color || "#217346";

            await context.sync();
            actionsDone++;
            counter.innerText = actionsDone;
        });

    } catch (e) {
        addMsg("Xatolik: " + e.message, 'ai');
    } finally {
        document.getElementById("status-dot").style.color = "#00ff00";
    }
}

function addMsg(t, c) {
    const d = document.createElement("div");
    d.className = `bubble ${c}`;
    d.innerText = t;
    chat.appendChild(d);
    chat.parentNode.scrollTop = chat.parentNode.scrollHeight;
}