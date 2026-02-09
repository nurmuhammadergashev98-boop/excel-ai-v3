Office.onReady();

const chat = document.getElementById("chat");
const btn = document.getElementById("sendBtn");
const input = document.getElementById("msg");

btn.onclick = async () => {
    const val = input.value.trim();
    if (!val) return;

    addMsg(val, 'user');
    input.value = "";
    
    const key = "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";
    
    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
            body: JSON.stringify({
                "model": "llama-3.3-70b-versatile",
                "messages": [{
                    "role": "system",
                    "content": "Sen Excel ustasisan. Foydalanuvchi buyrug'iga qarab JSON qaytar: {'reply': 'javob', 'action': 'write/format/formula', 'cell': 'A1', 'val': 'qiymat', 'color': 'hex_color'}"
                }, { "role": "user", "content": val }],
                "response_format": { "type": "json_object" }
            })
        });

        const data = await response.json();
        const res = JSON.parse(data.choices[0].message.content);
        
        addMsg(res.reply, 'ai');

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            let range = res.cell ? sheet.getRange(res.cell) : context.workbook.getSelectedRange();

            if (res.action === "write") range.values = [[res.val]];
            if (res.action === "formula") range.formulas = [[res.val]];
            if (res.action === "format") range.format.fill.color = res.color || "yellow";

            await context.sync();
        });
    } catch (err) {
        addMsg("Tizimda xatolik: " + err.message, 'ai');
    }
};

function addMsg(text, type) {
    const d = document.createElement("div");
    d.className = `bubble ${type}`;
    d.innerText = text;
    chat.appendChild(d);
    chat.scrollTop = chat.scrollHeight;
}