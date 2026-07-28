import express from 'express';
import cors from 'cors';
import { OpenAI } from 'openai';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import pptxgen from 'pptxgenjs';
import fs from 'fs';
import https from 'https';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// .env faylini yuklaymiz
dotenv.config({ path: path.resolve(__dirname, '../.env') });

const app = express();
app.use(cors({ origin: '*' }));
app.use(express.json());

console.log('BOT_TOKEN mavjud:', !!process.env.BOT_TOKEN);
console.log('MEGALLM_API_KEY mavjud:', !!process.env.MEGALLM_API_KEY);

// MegaLLM Client
const client = new OpenAI({
    baseURL: "https://ai.megallm.io/v1",
    apiKey: process.env.MEGALLM_API_KEY
});

const BOT_TOKEN = process.env.BOT_TOKEN;

// Telegram API ga fayl yuborish (native https, form-data siz)
async function sendPptxToTelegram(chatId, filePath, filename) {
    return new Promise((resolve, reject) => {
        const fileBuffer = fs.readFileSync(filePath);
        const boundary = '----FormBoundary' + Date.now();
        
        // Multipart form-data qo'lda yasaymiz
        const head = Buffer.from(
            `--${boundary}\r\n` +
            `Content-Disposition: form-data; name="chat_id"\r\n\r\n` +
            `${chatId}\r\n` +
            `--${boundary}\r\n` +
            `Content-Disposition: form-data; name="document"; filename="${filename}"\r\n` +
            `Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n`
        );
        const tail = Buffer.from(`\r\n--${boundary}--\r\n`);
        const body = Buffer.concat([head, fileBuffer, tail]);

        const options = {
            hostname: 'api.telegram.org',
            path: `/bot${BOT_TOKEN}/sendDocument`,
            method: 'POST',
            family: 4,
            headers: {
                'Content-Type': `multipart/form-data; boundary=${boundary}`,
                'Content-Length': body.length
            }
        };

        const req = https.request(options, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                try {
                    const parsed = JSON.parse(data);
                    if (parsed.ok) resolve(parsed);
                    else reject(new Error(`Telegram xatosi: ${parsed.description}`));
                } catch(e) {
                    reject(new Error('Telegram javobini parse qilib bo\'lmadi'));
                }
            });
        });

        req.on('error', reject);
        req.write(body);
        req.end();
    });
}

// Ana endpoint
app.post('/api/generate-slides', async (req, res) => {
    const { topic, chatId } = req.body;
    console.log(`\n📥 So'rov: topic="${topic}", chatId="${chatId}"`);

    if (!topic) return res.status(400).json({ success: false, error: "Mavzu kiritilmagan" });
    if (!chatId) return res.status(400).json({ success: false, error: "chatId topilmadi" });

    let fileName = null;

    try {
        // 1. AI orqali slayd kontent yaratish
        console.log('🤖 AI dan kontent so\'ralyapti...');
        const response = await client.chat.completions.create({
            model: "deepseek-ai/deepseek-v3.1",
            messages: [
                {
                    role: "system",
                    content: "Sen dars taqdimotlari uchun faqat qat'iy JSON formatida javob beradigan yordamchisan. Hech qanday markdown teglari yoki tushuntirish qo'shma."
                },
                {
                    role: "user",
                    content: `Mavzu: "${topic}". 5 ta slayd uchun qisqa kontent tayyorla. Har bir punkt maksimal 5-7 ta so'zdan iborat bo'lsin. Javob faqat shu formatda bo'lsin: [{"title": "Sarlavha", "bullets": ["1-punkt", "2-punkt"]}]`
                }
            ],
            max_tokens: 800
        });

        let rawContent = response.choices[0].message.content.trim();
        console.log('✅ AI javobi keldi, parse qilinmoqda...');
        
        if (rawContent.startsWith("```")) {
            rawContent = rawContent.replace(/^```json\s*/i, "").replace(/```$/, "").trim();
        }

        const slidesData = JSON.parse(rawContent);
        console.log(`✅ ${slidesData.length} ta slayd kontent tayyor`);

        // 2. PPTX fayl yaratish
        console.log('🎨 PPTX yaratilmoqda...');
        let pptx = new pptxgen();
        pptx.layout = 'LAYOUT_16x9';

        slidesData.forEach((slideData, i) => {
            let slide = pptx.addSlide();
            slide.background = { fill: "F8F9FA" };

            slide.addShape(pptx.ShapeType.rect, {
                x: 0, y: 6.8, w: "100%", h: 0.7,
                fill: { color: "1E3A8A" }
            });

            slide.addText(slideData.title || `Slayd ${i+1}`, {
                x: 0.8, y: 0.5, w: "85%", h: 0.8,
                fontSize: 24, bold: true, color: "1E3A8A",
                fontFace: "Arial"
            });

            if (slideData.bullets && slideData.bullets.length > 0) {
                const bodyText = slideData.bullets.join('\n\n');
                slide.addText(bodyText, {
                    x: 0.8, y: 1.6, w: "85%", h: 4.8,
                    fontSize: 16, color: "334155",
                    bullet: { type: 'number' },
                    fontFace: "Arial",
                    lineSpacing: 22
                });
            }
        });

        // 3. Faylni saqlash (backend papkasida)
        fileName = path.join(__dirname, `${Date.now()}_taqdimot.pptx`);
        await pptx.writeFile({ fileName });
        console.log(`✅ PPTX saqlandi: ${fileName}`);

        // 4. Telegram'ga yuborish
        console.log(`📤 Telegram'ga yuborilmoqda (chatId: ${chatId})...`);
        const friendlyName = `${topic.replace(/[^a-zA-Z0-9\s]/g, '').trim().replace(/\s+/g, '_').slice(0, 30)}_taqdimot.pptx`;
        await sendPptxToTelegram(chatId, fileName, friendlyName);
        console.log('✅ Telegram\'ga yuborildi!');

        // 5. Vaqtinchalik faylni o'chirish
        fs.unlinkSync(fileName);

        res.json({ success: true, slideCount: slidesData.length, slides: slidesData });

    } catch (error) {
        console.error("❌ Backend xatolik:", error.message || error);
        if (fileName && fs.existsSync(fileName)) {
            try { fs.unlinkSync(fileName); } catch(e) {}
        }
        res.status(500).json({ success: false, error: error.message || 'Noma\'lum xatolik' });
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n✅ Backend server localhost:${PORT} da ishga tushdi!\n`);
});