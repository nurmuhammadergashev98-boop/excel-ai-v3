Office.onReady();

const chatContainer = document.getElementById("chatContainer");
const promptInput = document.getElementById("prompt");
const sendBtn = document.getElementById("sendBtn");
const status = document.getElementById("status");

promptInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSend();
    }
});

sendBtn.onclick = handleSend;

async function handleSend() {
    const text = promptInput.value.trim();
    if (!text) return;

    addMessage(text, 'user');
    promptInput.value = "";
    status.innerText = "Nur AI bajarmoqda...";

    const apiKey = "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";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
            body: JSON.stringify({
                "model": "llama-3.3-70b-versatile",
                "messages": [{
                    "role": "system",
                    "content": `Sen Nurmuhammadning do'stona va professional Excel yordamchisisan.
                    Faqat JSON formatida javob ber. 
                    
                    Vazifalar:
                    1. 'write': Oddiy matn yoki jadvallar uchun (values ishlatiladi).
                    2. 'formula': Excel formulalari uchun (formulas ishlatiladi).
                    3. 'format': Kataklarni bo'yash uchun (color ishlatiladi).
                    
                    Qoidalar:
                    - Jadval yoki formula bo'lsa, 'data' massivi har doim [[...]] ko'rinishida bo'lsin.
                    - Agar foydalanuvchi 'A1 va B1 ni qo'sh' desa, action: 'formula', data: [['=A1+B1']] bo'ladi.
                    
                    Namuna:
                    {"reply": "Formula qo'yildi!", "action": "formula", "cell": "C1", "data": [["=SUM(A1:B1)"]]}`
                }, { "role": "user", "content": text }],
                "response_format": { "type": "json_object" }
            })
        });

        const data = await response.json();
        const res = JSON.parse(data.choices[0].message.content);

        addMessage(res.reply, 'ai');

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const rangeAddress = res.cell || "A1";
            const range = sheet.getRange(rangeAddress);

            if (res.action === "write") {
                range.values = res.data;
            } 
            else if (res.action === "formula") {
                range.formulas = res.data;
            } 
            else if (res.action === "format") {
                range.format.fill.color = res.color || "yellow";
            }

            await context.sync();
        });

    } catch (err) {
        addMessage("Xatolik: Buyruqni tushunishda xato bo'ldi.", 'ai');
        console.error(err);
    } finally {
        status.innerText = "online";
    }
}

function addMessage(text, type) {
    const div = document.createElement("div");
    div.className = `msg ${type}`;
    div.innerText = text;
    chatContainer.appendChild(div);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}