Office.onReady();

const chat = document.getElementById("chat-container");
const input = document.getElementById("prompt");
const btn = document.getElementById("sendBtn");
const counter = document.getElementById("counter");
let actionsDone = 0;

input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); run(); } });
btn.onclick = run;

async function run() {
    const text = input.value.trim();
    if (!text) return;

    addMsg(text, 'user');
    input.value = "";
    document.getElementById("status-dot").style.color = "orange";

    const key = "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";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
            body: JSON.stringify({
                "model": "llama-3.3-70b-versatile",
                "messages": [{
                    "role": "system",
                    "content": `Sen professional Excel avtomatizatorisan. Foydalanuvchi buyrug'iga ko'ra JSON qaytar.
                    Agarda jadval tuzish so'ralsa, 'data' qismiga massivlar massivini yoz (masalan: [[r1c1, r1c2], [r2c1, r2c2]]).
                    Agarda formula bo'lsa, uni '=' bilan boshla.
                    Struktura: {"reply": "insoniy javob", "action": "write/formula/format", "cell": "A1:B10", "data": [[...]], "color": "#hex"}`
                }, { "role": "user", "content": text }],
                "response_format": { "type": "json_object" }
            })
        });

        const json = await response.json();
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
        addMsg("Tizimda xatolik: " + e.message, 'ai');
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