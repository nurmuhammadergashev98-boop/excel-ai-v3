// ... (Office.onReady va boshqa o'zgaruvchilar o'sha-o'sha qoladi)

async function run() {
    const text = input.value.trim();
    const key = keyInput.value.trim();

    if (!text || !key) return;

    localStorage.setItem("groq_key", key);
    addMsg(text, 'user');
    input.value = "";
    document.getElementById("status-dot").style.color = "orange";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
            body: JSON.stringify({
                "model": "llama-3.3-70b-versatile",
                "messages": [{
                    "role": "system",
                    "content": `Siz professional Excel yordamchisiz. Foydalanuvchi buyrug'ini tahlil qiling va FAQAT JSON qaytaring.
                    
                    MUHIM QOIDALAR:
                    1. Agar foydalanuvchi jadval so'rasa, 'data' massivlar massivi bo'lsin: [[ustun1, ustun2], [qiymat1, qiymat2]].
                    2. 'cell' har doim ma'lumot boshlanadigan bitta katak bo'lsin (masalan: "A1").
                    3. JSON strukturasi: {"reply": "qisqa izoh", "action": "write/formula/format", "cell": "A1", "data": [[]], "color": "#hex"}`
                }, { "role": "user", "content": text }],
                "response_format": { "type": "json_object" }
            })
        });

        const json = await response.json();
        const res = JSON.parse(json.choices[0].message.content);
        addMsg(res.reply, 'ai');

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            if (res.action === "write" && res.data) {
                // DIQQAT: Mana shu qism o'lcham xatosini yo'qotadi
                const rowCount = res.data.length;
                const colCount = res.data[0].length;
                
                // Tanlangan katakdan boshlab, ma'lumot o'lchamiga ko'ra joyni kengaytiramiz
                const startRange = sheet.getRange(res.cell || "A1");
                const targetRange = startRange.getResizedRange(rowCount - 1, colCount - 1);
                
                targetRange.values = res.data;
                targetRange.format.autofitColumns(); // Ustunlarni chiroyli tekislaydi
            } 
            
            if (res.action === "formula") {
                sheet.getRange(res.cell || "A1").formulas = res.data;
            }
            
            if (res.action === "format") {
                sheet.getRange(res.cell || "A1").format.fill.color = res.color || "#217346";
            }

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