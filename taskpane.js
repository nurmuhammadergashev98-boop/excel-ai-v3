// taskpane.js ichidagi fetch qismidagi 'messages' bo'limini shunga almashtiring:

"messages": [
    { 
        "role": "system", 
        "content": `Siz professional Excel mutaxassisiz. 
        1. FAQAT O'ZBEK TILIDA javob bering. 
        2. Har doim JSON qaytaring. 
        3. Agar foydalanuvchi "o'chir" yoki "tozala" desa, action: "clear" deb belgilang.
        4. Strukturani buzmang: {"reply": "...", "action": "write/clear/format", "cell": "A1", "data": [[]]}` 
    },
    { "role": "user", "content": text }
]

// Pastroqdagi Excel.run qismini esa mana shunday boyiting:

await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    if (res.action === "write" && res.data) {
        const targetRange = sheet.getRange(res.cell || "A1").getResizedRange(res.data.length - 1, res.data[0].length - 1);
        targetRange.values = res.data;
    } 
    else if (res.action === "clear") {
        // Hamma ishlatilgan kataklarni topib tozalaydi
        sheet.getUsedRange().clear();
        addMsg("Varaq butunlay tozalandi.", "ai");
    }
    else if (res.action === "format") {
        sheet.getRange(res.cell || "A1").format.fill.color = res.color || "yellow";
    }

    await context.sync();
});