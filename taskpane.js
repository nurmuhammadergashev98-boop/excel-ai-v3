Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Excel muhiti tayyor.");
    }
    // Tugma bosilishini har doim eshitib turadi
    document.getElementById("run").onclick = callGroqAI;
});

async function callGroqAI() {
    const prompt = document.getElementById("prompt").value;
    const status = document.getElementById("status");

    if (!prompt) {
        alert("Iltimos, vazifani yozing!");
        return;
    }

    status.style.display = "block";
    status.innerText = "AI o'ylamoqda...";

    // API Kalit (Shifrlanmagan holda, tekshirish oson bo'lishi uchun)
    const apiKey = "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";

    try {
        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + apiKey
            },
            body: JSON.stringify({
                model: "mixtral-8x7b-32768",
                messages: [
                    { role: "system", content: "Sen Excel mutaxassisisan. Faqat formula yoki qisqa javob qaytar." },
                    { role: "user", content: prompt }
                ]
            })
        });

        if (!response.ok) {
            throw new Error("API xatosi: " + response.status);
        }

        const data = await response.json();
        const aiResponse = data.choices[0].message.content.trim();

        // 1. Birinchi bo'lib natijani ekranda ko'rsatamiz (Brauzerda ham ko'rinadi)
        status.innerText = "AI Javobi: " + aiResponse;
        
        // 2. Agar Excel ichida bo'lsak, katakka yozamiz
        if (typeof Excel !== 'undefined' && Office.context.host === Office.HostType.Excel) {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                if (aiResponse.startsWith("=")) {
                    range.formulas = [[aiResponse]];
                } else {
                    range.values = [[aiResponse]];
                }
                await context.sync();
            });
        } else {
            // Agar Excel tashqarisida (brauzerda) bo'lsak, ogohlantirish chiqaramiz
            alert("AI javob berdi: " + aiResponse + "\n(Excel topilmadi, faqat natija ko'rsatildi)");
        }

    } catch (error) {
        status.innerText = "Xatolik: " + error.message;
        console.error(error);
        alert("Xatolik yuz berdi: " + error.message);
    }
}