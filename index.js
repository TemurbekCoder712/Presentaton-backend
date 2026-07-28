import express from 'express';
import cors from 'cors';
import { OpenAI } from 'openai';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import pptxgen from 'pptxgenjs';
import fs from 'fs';
import https from 'https';
import swaggerUi from 'swagger-ui-express';
import swaggerJsdoc from 'swagger-jsdoc';
import prisma from './prismaClient.js';
import { GoogleGenerativeAI } from '@google/generative-ai';

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

// Gemini Client
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

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

// Swagger sozlamalari
const swaggerOptions = {
  definition: {
    openapi: '3.0.0',
    info: {
      title: 'Presentation AI API',
      version: '1.0.0',
      description: 'Telegram bot va Web App uchun taqdimot yaratish API',
    },
    servers: [
      {
        url: 'https://presentaton-backend.onrender.com',
        description: 'Render Server',
      },
      {
        url: 'http://localhost:5000',
        description: 'Local Server',
      }
    ],
  },
  apis: [fileURLToPath(import.meta.url)], // o'z faylini o'qiydi
};

const swaggerSpec = swaggerJsdoc(swaggerOptions);
app.use('/api-docs', swaggerUi.serve, swaggerUi.setup(swaggerSpec));

/**
 * @swagger
 * /api/generate-slides:
 *   post:
 *     summary: Sun'iy intellekt orqali taqdimot (PPTX) yaratish
 *     description: Mavzu va Telegram Chat ID qabul qilib, AI orqali PPTX yaratadi va bot orqali foydalanuvchiga yuboradi.
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               topic:
 *                 type: string
 *                 example: "Fizika yorug'lik tezligi"
 *               chatId:
 *                 type: string
 *                 example: "123456789"
 *     responses:
 *       200:
 *         description: Muvaffaqiyatli yaratildi va jo'natildi
 *       400:
 *         description: Noto'g'ri so'rov (mavzu yoki chatId yo'q)
 *       500:
 *         description: Server xatosi
 */
app.post('/api/generate-slides', async (req, res) => {
    const { topic, chatId } = req.body;
    console.log(`\n📥 So'rov: topic="${topic}", chatId="${chatId}"`);

    if (!topic) return res.status(400).json({ success: false, error: "Mavzu kiritilmagan" });
    if (!chatId) return res.status(400).json({ success: false, error: "chatId topilmadi" });

    let fileName = null;

    try {
        // Balansni tekshiramiz
        // Balansni tekshiramiz (MVP uchun o'chirildi)
        /*
        const user = await prisma.user.findUnique({
            where: { telegramId: BigInt(chatId) }
        });

        if (!user) {
            return res.status(400).json({ success: false, error: "Foydalanuvchi topilmadi. Iltimos, Telegram botga kirib /start tugmasini bosing." });
        }

        if (user.balance <= 0) {
            return res.status(403).json({ success: false, error: "Sizning balansingiz (limit) tugagan. Iltimos, hisobingizni to'ldiring!" });
        }
        */

        // 1. AI orqali slayd kontent yaratish (MegaLLM - Miya)
        console.log('🧠 MegaLLM dan kontent so\'ralyapti...');
        const megaResponse = await client.chat.completions.create({
            model: "deepseek-ai/deepseek-v3.1",
            messages: [
                {
                    role: "system",
                    content: "Sen eng aqlli o'qituvchisan. Berilgan mavzu bo'yicha eng muhim va qiziqarli faktlarni yig'ib ber. 5 ta slayd uchun reja va qisqacha ma'lumotlar yoz. Hech qanday markdown ishlatma."
                },
                {
                    role: "user",
                    content: `Mavzu: "${topic}"`
                }
            ],
            max_tokens: 1500
        });

        const rawContent = megaResponse.choices[0].message.content.trim();
        console.log('✅ MegaLLM javob berdi. Gemini dizaynga o\'tkazmoqda...');

        // 2. Gemini orqali Dizayn va Formatlash (Gemini - Dizayner)
        const geminiModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const geminiPrompt = `Senga xom matn beraman. Sen uni PPTX dizayni uchun qat'iy JSON array formatiga o'tkazishing kerak. Har bir slayd uchun quyidagi maydonlar bo'lishi shart:
- title: Slayd sarlavhasi
- bullets: Slayddagi matnlar (qisqa qisqa punktlar arrayi)
- layout: "split-right" (o'ngda rasm), "split-left" (chapda rasm), yoki "centered" (faqat matn)
- themeColor: Mavzuga mos bitta HEX rang kodi (faqat 6 ta harf/raqam, masalan: 1E3A8A)
- keyword: Rasm qidirish uchun mos inglizcha bitta so'z (masalan: "space", "computer")

Faqat JSON formatda qaytar, boshqa hech qanday izoh qo'shma.
Matn:
${rawContent}`;

        const geminiResult = await geminiModel.generateContent(geminiPrompt);
        let geminiJsonStr = geminiResult.response.text().trim();
        
        if (geminiJsonStr.startsWith("```")) {
            geminiJsonStr = geminiJsonStr.replace(/^```json\s*/i, "").replace(/```$/, "").trim();
        }

        const slidesData = JSON.parse(geminiJsonStr);
        console.log(`✅ Gemini ${slidesData.length} ta dizaynli slayd tayyorladi.`);

        // 3. PPTX fayl yaratish
        console.log('🎨 PPTX yaratilmoqda...');
        let pptx = new pptxgen();
        pptx.layout = 'LAYOUT_16x9';

        slidesData.forEach((slideData, i) => {
            let slide = pptx.addSlide();
            slide.background = { fill: "F8F9FA" };

            const themeColor = (slideData.themeColor || "1E3A8A").replace('#', '');
            const layout = slideData.layout || "split-right";

            // Dizayn dekoratsiyasi (yuqori yoki pastda chiziq)
            slide.addShape(pptx.ShapeType.rect, {
                x: 0, y: 6.8, w: "100%", h: 0.7,
                fill: { color: themeColor }
            });

            if (layout === "split-right") {
                slide.addText(slideData.title || `Slayd ${i+1}`, {
                    x: 0.5, y: 0.5, w: "50%", h: 0.8,
                    fontSize: 26, bold: true, color: themeColor, fontFace: "Arial"
                });
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join('\n\n'), {
                        x: 0.5, y: 1.5, w: "50%", h: 4.8,
                        fontSize: 18, color: "334155", bullet: { type: 'number' }, fontFace: "Arial", lineSpacing: 22
                    });
                }
            } else if (layout === "split-left") {
                slide.addText(slideData.title || `Slayd ${i+1}`, {
                    x: 4.5, y: 0.5, w: "50%", h: 0.8,
                    fontSize: 26, bold: true, color: themeColor, fontFace: "Arial"
                });
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join('\n\n'), {
                        x: 4.5, y: 1.5, w: "50%", h: 4.8,
                        fontSize: 18, color: "334155", bullet: { type: 'number' }, fontFace: "Arial", lineSpacing: 22
                    });
                }
            } else {
                slide.addText(slideData.title || `Slayd ${i+1}`, {
                    x: 1.0, y: 0.5, w: "80%", h: 0.8, align: "center",
                    fontSize: 28, bold: true, color: themeColor, fontFace: "Arial"
                });
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join('\n\n'), {
                        x: 1.0, y: 1.5, w: "80%", h: 4.8,
                        fontSize: 20, color: "334155", align: "center", fontFace: "Arial", lineSpacing: 24
                    });
                }
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

        // 6. Balansni ayirish (MVP uchun o'chirildi)
        /*
        await prisma.user.update({
            where: { telegramId: BigInt(chatId) },
            data: { balance: { decrement: 1 } }
        });
        */

        res.json({ success: true, slideCount: slidesData.length, slides: slidesData });

    } catch (error) {
        console.error("❌ Backend xatolik:", error.message || error);
        if (fileName && fs.existsSync(fileName)) {
            try { fs.unlinkSync(fileName); } catch(e) {}
        }
        res.status(500).json({ success: false, error: error.message || 'Noma\'lum xatolik' });
    }
});

/**
 * @swagger
 * /api/support-chat:
 *   post:
 *     summary: AI Yordamchi bilan suhbatlashish
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               message:
 *                 type: string
 *     responses:
 *       200:
 *         description: AI javobi
 */
app.post('/api/support-chat', async (req, res) => {
    const { message } = req.body;
    if (!message) return res.status(400).json({ success: false, error: "Xabar kiritilmagan" });

    try {
        const geminiModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const systemPrompt = "Sen Presentation AI loyihasining rasmiy maslahatchisi (Support Bot) san. Yaratuvching: Temurbek (TemurbekCoder). Sen faqat loyiha qanday ishlashi, narxlar va taqdimotlar tayyorlash haqida yordam berasan. Qisqa, samimiy va o'zbek tilida javob ber. Agar dasturlash, matematika, siyosat yoki boshqa umuman aloqasi yo'q mavzuda savol berishsa, uzr so'rab o'z ishingga qayt.";
        
        const result = await geminiModel.generateContent(`${systemPrompt}\n\nFoydalanuvchi: ${message}`);
        const reply = result.response.text();

        res.json({ success: true, reply });
    } catch (err) {
        console.error("Support Chat xatosi:", err);
        res.status(500).json({ success: false, error: "Server xatosi yuz berdi" });
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n✅ Backend server localhost:${PORT} da ishga tushdi!\n`);
});

// Telegram botni ishga tushirish
import './bot.js';