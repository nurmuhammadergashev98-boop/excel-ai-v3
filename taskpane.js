Office.onReady();

const chatContainer = document.getElementById("chatContainer");
const promptInput = document.getElementById("prompt");
const sendBtn = document.getElementById("sendBtn");
const status = document.getElementById("status");

// ENTER TUGMASINI BOSGANDA YUBORISH
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
    status.innerText = "typing...";

    const apiKey = "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";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + apiKey },
            body: JSON.stringify({
                "model": "llama-3.3-70b-versatile",
                "messages": [{
                    "role": "system",
                    "content": "Sen Nurmuhammadning do'stona Excel yordamchisisan. Suhbat uslubing samimiy bo'lsin. Excel buyruqlari kelsa, faqat JSON qaytar: {'reply': 'javob', 'action': 'write/format/formula/none', 'cell': 'A1', 'val': '...', 'color': 'hex'}"
                }, { "role": "user", "content": text }],
                "response_format": { "type": "json_object" }
            })
        });

        const data = await response.json();
        const res = JSON.parse(data.choices[0].message.content);

        addMessage(res.reply, 'ai');

        if (res.action !== "none") {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                let range = res.cell ? sheet.getRange(res.cell) : context.workbook.getSelectedRange();

                if (res.action === "write") range.values = [[res.val]];
                if (res.action === "formula") range.formulas = [[res.val]];
                if (res.action === "format") range.format.fill.color = res.color || "#FFFF00";

                await context.sync();
            });
        }
    } catch (err) {
        addMessage("Xatolik: " + err.message, 'ai');
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