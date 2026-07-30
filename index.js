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
import { appendFeedbackToSheet, updateUserStatus } from './googleSheets.js';
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

async function sendTextMessageToTelegram(chatId, text) {
    try {
        await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'HTML' })
        });
    } catch (e) {
        console.error("Telegramga xabar yuborishda xatolik:", e);
    }
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
/**
 * @swagger
 * /api/auth:
 *   post:
 *     summary: Telegram foydalanuvchisini ro'yxatdan o'tkazish
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               id:
 *                 type: number
 *               first_name:
 *                 type: string
 *               last_name:
 *                 type: string
 *               username:
 *                 type: string
 *     responses:
 *       200:
 *         description: Foydalanuvchi ma'lumotlari
 */
app.post('/api/auth', async (req, res) => {
    const { id, first_name, last_name, username } = req.body;
    if (!id) return res.status(400).json({ success: false, error: "Telegram ID mavjud emas" });

    try {
        const user = await prisma.user.upsert({
            where: { telegramId: BigInt(id) },
            update: {
                firstName: first_name || "Foydalanuvchi",
                lastName: last_name || null,
                username: username || null,
            },
            create: {
                telegramId: BigInt(id),
                firstName: first_name || "Foydalanuvchi",
                lastName: last_name || null,
                username: username || null,
                balance: 1,
                planType: 'free'
            }
        });
        
        // Prisma returns BigInt which JSON.stringify cannot serialize. So convert to string.
        
        // Fonda holatni yangilash (WebApp ochilganda)
        updateUserStatus(id).catch(e => console.error("Holat yangilashda xato (auth):", e));

        res.json({
            success: true,
            user: {
                ...user,
                telegramId: user.telegramId.toString(),
                createdAt: user.createdAt.toISOString()
            }
        });
    } catch (err) {
        console.error("Auth xatosi:", err);
        res.status(500).json({ success: false, error: "Bazaga yozishda xatolik" });
    }
});

app.post('/api/feedback', async (req, res) => {
    const { name, username, feedback } = req.body;
    
    if (!feedback) {
        return res.status(400).json({ success: false, error: "Fikr kiritilmagan" });
    }

    try {
        // 1. Google Sheets ga yozish
        await appendFeedbackToSheet(name, username, feedback);

        // 2. Telegram orqali adminga xabar berish
        const ADMIN_CHAT_ID = "6150067773";
        const message = `🔔 <b>Yangi fikr/taklif keldi!</b>\n\n👤 <b>Ism:</b> ${name || 'Noma\'lum'}\n🔗 <b>Username:</b> ${username ? '@'+username.replace('@','') : 'Yo\'q'}\n\n💬 <b>Fikr:</b>\n<i>${feedback}</i>`;
        
        await sendTextMessageToTelegram(ADMIN_CHAT_ID, message);

        res.json({ success: true, message: "Fikringiz qabul qilindi!" });
    } catch (e) {
        console.error("Feedback xatosi:", e);
        res.status(500).json({ success: false, error: "Tizim xatosi" });
    }
});

app.post('/api/generate-slides', async (req, res) => {
    const { topic, chatId } = req.body;
    console.log(`\n📥 So'rov: topic="${topic}", chatId="${chatId}"`);

    if (!topic) return res.status(400).json({ success: false, error: "Mavzu kiritilmagan" });
    if (!chatId) return res.status(400).json({ success: false, error: "chatId topilmadi" });

    // Holatni va mavzuni fonda yangilash
    updateUserStatus(chatId, topic).catch(e => console.error("Holatni yangilashda xato:", e));

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
        const geminiModel = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
        const geminiPrompt = `Senga xom matn beraman. Sen uni PPTX dizayni uchun qat'iy JSON formatiga o'tkazishing kerak. Har bir slayd uchun quyidagi maydonlar bo'lishi shart:
{
  "themeColor": "1E3A8A", // Mavzuga mos bitta asosiy rang (masalan: Iqtisodiyot uchun 1E3A8A)
  "slides": [
    {
      "title": "Slayd sarlavhasi",
      "bullets": ["1-punkt", "2-punkt"], // Qisqa matnlar
      "layout": "split-right", // yoki "split-left" yoki "centered"
      "keyword": "inglizcha rasm qidiruv so'zi",
      "chart": { // Agar mavzuga mos statistika bo'lsa, uni qo'sh. Aks holda null. (Faqat 1-2 ta slaydda)
         "type": "bar", // yoki "pie"
         "title": "Grafika nomi",
         "labels": ["A", "B", "C"],
         "values": [10, 20, 30]
      }
    }
  ]
}
Faqat va faqat JSON obyekt qaytar, boshqa hech qanday so'z qo'shma.
Matn:
${rawContent}`;

        const geminiResult = await geminiModel.generateContent(geminiPrompt);
        let geminiJsonStr = geminiResult.response.text().trim();
        
        if (geminiJsonStr.startsWith("```")) {
            geminiJsonStr = geminiJsonStr.replace(/^```json\s*/i, "").replace(/```$/, "").trim();
        }

        let presentationData;
        try {
            presentationData = JSON.parse(geminiJsonStr);
        } catch (e) {
            let temp = JSON.parse(`[${geminiJsonStr.split('[')[1].split(']')[0]}]`);
            presentationData = { themeColor: "6c63ff", slides: temp };
        }
        if (Array.isArray(presentationData)) {
            presentationData = { themeColor: "6c63ff", slides: presentationData };
        }
        const slidesData = presentationData.slides || [];
        const globalThemeColor = (presentationData.themeColor || "6c63ff").replace('#', '');
        console.log(`✅ Gemini ${slidesData.length} ta dizaynli slayd tayyorladi.`);

        // 3. PPTX fayl yaratish
        console.log('🎨 PPTX yaratilmoqda...');
        let pptx = new pptxgen();
        pptx.layout = 'LAYOUT_16x9';

        async function downloadImageLocal(url, filepath) {
            try {
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 2000); // 2 sekund kutish
                
                const res = await fetch(url, { signal: controller.signal });
                clearTimeout(timeoutId);
                
                if (!res.ok) throw new Error();
                const arrayBuffer = await res.arrayBuffer();
                fs.writeFileSync(filepath, Buffer.from(arrayBuffer));
                return true;
            } catch(e) { return false; }
        }

        for (let i = 0; i < slidesData.length; i++) {
            const slideData = slidesData[i];
            let slide = pptx.addSlide();
            const themeColor = globalThemeColor;
            const layout = slideData.layout || "centered";
            
            // 1) Asosiy (Sarlavha) Slayd dizayni
            if (i === 0) {
                slide.background = { fill: themeColor }; 
                slide.addShape(pptx.ShapeType.rect, { x: -1, y: -1, w: 4, h: 4, fill: { color: "FFFFFF", transparency: 85 } });
                slide.addShape(pptx.ShapeType.rect, { x: 7.5, y: 4.5, w: 5, h: 5, fill: { color: "FFFFFF", transparency: 90 } });

                slide.addText(slideData.title || `Slayd ${i+1}`, {
                    x: 0.5, y: 2.0, w: 9, h: 1.5, align: "center",
                    fontSize: 44, bold: true, color: "FFFFFF", fontFace: "Arial", wrap: true
                });
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join(' | '), {
                        x: 0.5, y: 3.5, w: 9, h: 1.5, align: "center",
                        fontSize: 18, color: "F8F9FA", fontFace: "Arial", italic: true, wrap: true
                    });
                }
                continue;
            }

            // 2) Qolgan slaydlar foni
            slide.background = { fill: "F4F5F8" };

            // Zamonaviy Header (Sarlavha qismi)
            slide.addShape(pptx.ShapeType.rect, {
                x: 0, y: 0, w: "100%", h: 1.2, fill: { color: themeColor }
            });
            slide.addText(slideData.title || `Slayd ${i+1}`, {
                x: 0.5, y: 0.1, w: "90%", h: 1.0,
                fontSize: 26, bold: true, color: "FFFFFF", fontFace: "Arial", align: "left", valign: "middle", wrap: true
            });

            // Pastki chiziq va raqamlash
            slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.3, w: "100%", h: 0.2, fill: { color: themeColor } });
            slide.addText(`${i}`, { x: 9.0, y: 7.0, w: 0.8, h: 0.3, align: "right", fontSize: 12, color: themeColor, fontFace: "Arial" });

            const hasChart = slideData.chart && slideData.chart.labels && slideData.chart.values && slideData.chart.labels.length > 0;

            // Layoutlar
            if (layout === "split-right" || (layout === "centered" && hasChart)) {
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join('\n\n'), {
                        x: 0.5, y: 1.5, w: 4.8, h: 5.2, fontSize: 16, color: "2D3748", fontFace: "Arial", bullet: true, valign: "top", wrap: true
                    });
                }
                if (hasChart) {
                    const chartData = [{ name: slideData.chart.title || 'Statistika', labels: slideData.chart.labels, values: slideData.chart.values }];
                    const chartType = slideData.chart.type === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
                    slide.addChart(chartType, chartData, {
                        x: 5.6, y: 1.5, w: 4.0, h: 4.5, showLegend: true, legendPos: 'b',
                        showTitle: true, title: slideData.chart.title, titleFontSize: 14, titleColor: "2D3748",
                        barDir: 'col', showValue: true, chartColors: [themeColor, "3B82F6", "10B981", "F59E0B"]
                    });
                } else {
                    const imgFilename = `temp_web_${Date.now()}_${i}.jpg`;
                    const ok = await downloadImageLocal(`https://loremflickr.com/600/450/${encodeURIComponent(slideData.keyword || "education")}`, imgFilename);
                    if(ok) {
                        slide.addImage({ path: imgFilename, x: 5.6, y: 1.5, w: 4.0, h: 4.5, sizing: { type: 'cover', w: 4.0, h: 4.5 } });
                        setTimeout(() => { try { fs.unlinkSync(imgFilename); } catch(e){} }, 5000);
                    } else {
                        slide.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.5, w: 4.0, h: 4.5, fill: { color: themeColor, transparency: 80 } });
                        slide.addText(slideData.keyword ? slideData.keyword.toUpperCase() : 'VISUAL', { x: 5.6, y: 3.5, w: 4.0, h: 1, align: "center", fontSize: 24, bold: true, color: themeColor });
                    }
                }
            } else if (layout === "split-left") {
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join('\n\n'), {
                        x: 4.8, y: 1.5, w: 4.8, h: 5.2, fontSize: 16, color: "2D3748", fontFace: "Arial", bullet: true, valign: "top", wrap: true
                    });
                }
                if (hasChart) {
                    const chartData = [{ name: slideData.chart.title || 'Statistika', labels: slideData.chart.labels, values: slideData.chart.values }];
                    const chartType = slideData.chart.type === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
                    slide.addChart(chartType, chartData, {
                        x: 0.4, y: 1.5, w: 4.0, h: 4.5, showLegend: true, legendPos: 'b',
                        showTitle: true, title: slideData.chart.title, titleFontSize: 14, titleColor: "2D3748",
                        barDir: 'col', showValue: true, chartColors: [themeColor, "3B82F6", "10B981", "F59E0B"]
                    });
                } else {
                    const imgFilename = `temp_web_${Date.now()}_${i}.jpg`;
                    const ok = await downloadImageLocal(`https://loremflickr.com/600/450/${encodeURIComponent(slideData.keyword || "education")}`, imgFilename);
                    if(ok) {
                        slide.addImage({ path: imgFilename, x: 0.4, y: 1.5, w: 4.0, h: 4.5, sizing: { type: 'cover', w: 4.0, h: 4.5 } });
                        setTimeout(() => { try { fs.unlinkSync(imgFilename); } catch(e){} }, 5000);
                    } else {
                        slide.addShape(pptx.ShapeType.rect, { x: 0.4, y: 1.5, w: 4.0, h: 4.5, fill: { color: themeColor, transparency: 80 } });
                        slide.addText(slideData.keyword ? slideData.keyword.toUpperCase() : 'INFO', { x: 0.4, y: 3.5, w: 4.0, h: 1, align: "center", fontSize: 24, bold: true, color: themeColor });
                    }
                }
            } else {
                // Centered - Karta uslubida
                slide.addShape(pptx.ShapeType.rect, {
                    x: 1.0, y: 1.5, w: 8.0, h: 5.2, fill: { color: "FFFFFF" }
                });
                if (slideData.bullets && slideData.bullets.length > 0) {
                    slide.addText(slideData.bullets.join('\n\n'), {
                        x: 1.5, y: 1.8, w: 7.0, h: 4.6, fontSize: 18, color: "2D3748", fontFace: "Arial", align: "center", valign: "middle", wrap: true
                    });
                }
            }
        }

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
    const { message, history } = req.body;
    if (!message) return res.status(400).json({ success: false, error: "Xabar kiritilmagan" });

    try {
        const geminiModel = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
        const systemPrompt = `Sen "Presentation AI" loyihasining rasmiy maslahatchisi (Support Bot) san. Yaratuvching: kuchli dasturchi Temurbek Narzullayev (TemurbekCoder).
Vazifang: loyiha qanday ishlashi, narxlar va taqdimotlar tayyorlash haqida batafsil yordam berish.
Qisqa emas, balki foydalanuvchiga tushunarli bo'lishi uchun yetarlicha to'liq ma'lumot ber. 

QATIY QOIDA (SALOMLASHISH): Agar foydalanuvchi birinchi marta yozayotgan bo'lsa yoki shunchaki "Salom" deb yozsa, albatta loyihaning nima qila olishini (sun'iy intellekt yordamida soniyalarda ajoyib taqdimotlar yasashini), qanday tariflar borligini qisqacha aytib o't va o'zingni yaratuvching Temurbek Narzullayev (kuchli dasturchi) ekanligini faxr bilan ta'kidla. Shundan so'ng, ularga qanday yordam bera olishingni so'ra.

Agar dasturlash, matematika, siyosat yoki loyihaga aloqasi yo'q mavzuda savol berishsa, uzr so'rab javob berishdan bosh tort.

QATIY QOIDA (FORMATLASH): Hech qanday Markdown formatidan foydalanma! Yulduzchalar (** qalin yozuv uchun) yoki boshqa belgilarni ishlatma. Faqat toza matn (plain text) va abzaslar (yangi qatorlar) dan foydalan.

NARXLAR (TARIFLAR):
1. Start: 3,000 UZS (2-3 ta taqdimot)
2. Ommabop (Pro): 8,000 UZS (7-8 ta taqdimot, eng ko'p olinadigan tarif)
3. Premium: 16,000 UZS (14-15 ta taqdimot)
4. Teacher: 50,000 UZS (50+ ta taqdimot)
5. School: 100,000 UZS (Cheksiz taqdimot, 1 oylik limit)

Foydalanuvchi qachon narxlar haqida so'rasa, faqat shu tariflarni oddiy matn ko'rinishida yozib ber.`;

        let historyText = "";
        if (history && Array.isArray(history)) {
            // Take the last 5 messages to save tokens and keep context fresh
            const recentHistory = history.slice(-5);
            historyText = "--- Avvalgi yozishmalar tarixi ---\n" + recentHistory.map(m => `${m.role === 'user' ? 'Foydalanuvchi' : 'Sen'}: ${m.text}`).join('\n') + "\n----------------------------------\n\n";
        }
        
        const result = await geminiModel.generateContent(`${systemPrompt}\n\n${historyText}Foydalanuvchi so'rovi: ${message}`);
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