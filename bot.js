import { Telegraf } from 'telegraf';
import { OpenAI } from 'openai';
import dotenv from 'dotenv';
import path from 'path';
import pptxgen from 'pptxgenjs';
import fs from 'fs';
import axios from 'axios';
import https from 'https';
import { fileURLToPath } from 'url';
import prisma from './prismaClient.js';
import { appendUserToSheet } from './googleSheets.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// .env faylini yuklaymiz
dotenv.config({ path: path.resolve(__dirname, '../.env') });

const token = process.env.TELEGRAM_BOT_TOKEN || process.env.BOT_TOKEN || process.env.TOKEN;

if (!token) {
    console.error("❌ Xatolik: .env faylida Telegram bot tokeni topilmadi!");
} else {
    // Telegraf — timeout va agent sozlamalari bilan (IPv4 majburiy)
    const bot = new Telegraf(token, {
        handlerTimeout: 90_000,
        telegram: {
            agent: new https.Agent({ keepAlive: true, family: 4 }),
            attachmentAgent: new https.Agent({ keepAlive: true, family: 4 })
        }
    });

    const client = new OpenAI({
        baseURL: "https://ai.megallm.io/v1",
        apiKey: process.env.MEGALLM_API_KEY
    });

    // Rasmni vaqtinchalik yuklab olish funksiyasi
    async function downloadImage(url, filepath) {
        const response = await axios({
            url,
            method: 'GET',
            responseType: 'stream'
        });
        return new Promise((resolve, reject) => {
            const writer = fs.createWriteStream(filepath);
            response.data.pipe(writer);
            writer.on('finish', resolve);
            writer.on('error', reject);
        });
    }

    // /start buyrug'i
    bot.start(async (ctx) => {
        const chatId = ctx.from.id;
        const firstName = ctx.from.first_name || 'Ustoz';

        try {
            /* 
            const existingUser = await prisma.user.findUnique({
                where: { telegramId: BigInt(chatId) }
            });

            if (!existingUser) {
                await prisma.user.create({
                    data: {
                        telegramId: BigInt(chatId),
                        firstName: firstName,
                        lastName: ctx.from.last_name || null,
                        username: ctx.from.username || null,
                        balance: 1, // 1 ta bepul
                        planType: "free"
                    }
                });
                console.log(`✅ Yangi foydalanuvchi bazaga qo'shildi: ${firstName} (${chatId})`);
            }
            */
        } catch (e) {
            console.error("❌ Baza bilan ulanishda xatolik:", e);
        }

        // Google Sheets ga foydalanuvchini yozib qo'yish
        appendUserToSheet(ctx.from);

        // Admin ga xabar yuborish
        try {
            const adminId = '6150067773';
            const dateStr = new Date().toLocaleString('ru-RU', { timeZone: 'Asia/Tashkent' });
            const msg = `🔔 <b>Yangi foydalanuvchi botga kirdi!</b>\n\n👤 <b>Ism:</b> ${ctx.from.first_name || 'Mavjud emas'}\n👥 <b>Familiya:</b> ${ctx.from.last_name || 'Mavjud emas'}\n🔗 <b>Username:</b> ${ctx.from.username ? '@' + ctx.from.username : 'Mavjud emas'}\n📅 <b>Sana:</b> ${dateStr}`;
            bot.telegram.sendMessage(adminId, msg, { parse_mode: 'HTML' }).catch(err => {
                console.error("Adminga xabar yuborishda xatolik:", err);
            });
        } catch (err) {
            console.error("Adminga xabar yuborish parametrlarida xatolik:", err);
        }

        const baseUrl = process.env.WEBAPP_URL || 'https://presentaton-frontend.vercel.app';
        const webAppUrl = `${baseUrl}?chatId=${chatId}`;

        ctx.reply(
            `👋 Salom, ${firstName}!\n\n` +
            `📊 <b>Presentation AI</b> — dars taqdimotlarini sekundlar ichida yaratib beradigan aqlli yordamchingiz.\n\n` +
            `✅ Faqat mavzuni yozing — qolganini AI qiladi:\n` +
            `• Professional slaydlar\n` +
            `• Mos kontent va strukturа\n` +
            `• Tayyor PPTX fayl — shu yerga yuboriladi\n\n` +
            `⬇️ Boshlash uchun tugmani bosing:`,
            {
                parse_mode: 'HTML',
                reply_markup: {
                    inline_keyboard: [
                        [{ text: "🚀 Taqdimot yaratish", web_app: { url: webAppUrl } }]
                    ]
                }
            }
        );
    });

    // Matnli xabarlarni ushlash
    bot.on('text', async (ctx) => {
        const topic = ctx.message.text.trim();
        if (topic.startsWith('/')) return;

        const loadingMessage = await ctx.reply('🔄 Sun\'iy intellekt slayd kontentini tayyorlamoqda...');

        try {
            const response = await client.chat.completions.create({
                model: "deepseek-ai/deepseek-v3.1",
                messages: [
                    {
                        role: "system",
                        content: "Sen taqdimot dizayneri va analitiksan. Faqat qat'iy JSON formatida javob ber. Hech qanday markdown teglari, izohlar yoki qo'shimcha so'zlar qo'shma."
                    },
                    {
                        role: "user",
                        content: `Mavzu: "${topic}". Ushbu mavzu bo'yicha 5 ta slayd uchun professional kontent tayyorla.
Javobing aynan mana shu JSON strukturada bo'lishi shart:
{
  "themeColor": "1E3A8A", // Mavzuga mos bitta asosiy rang (masalan, Iqtisodiyot uchun 1E3A8A, Tabiat uchun 2E8B57, va hokazo)
  "slides": [
    {
      "title": "Slayd sarlavhasi",
      "bullets": ["1-qisqa qoida", "2-fakt"], // Har bir punkt qisqa va aniq bo'lsin
      "keyword": "inglizcha rasm qidiruv so'zi",
      "chart": { // Agar shu slaydga statistika kerak deb hisoblasang, qo'sh. Aks holda null qo'y. (Maksimal 1 yoki 2 ta slaydda chart bo'lsin)
         "type": "bar", // yoki "pie"
         "title": "Grafika nomi",
         "labels": ["A", "B", "C"],
         "values": [10, 20, 30]
      }
    }
  ]
}
Faqat toza JSON formatni qaytar.`
                    }
                ],
                max_tokens: 1500
            });

            let rawContent = response.choices[0].message.content.trim();
            if (rawContent.startsWith("```")) {
                rawContent = rawContent.replace(/^```json\s*/i, "").replace(/```$/, "").trim();
            }

            let presentationData;
            try {
                presentationData = JSON.parse(rawContent);
            } catch (e) {
                console.error("JSON parse error:", e);
                throw new Error("AI noto'g'ri formatda ma'lumot qaytardi.");
            }

            // Agar AI faqat array qaytarsa, uni to'g'rilash
            if (Array.isArray(presentationData)) {
                presentationData = { themeColor: "1E3A8A", slides: presentationData };
            }
            const slidesData = presentationData.slides || [];
            const themeColor = (presentationData.themeColor || "1E3A8A").replace('#', '');

            await ctx.telegram.editMessageText(
                ctx.chat.id, 
                loadingMessage.message_id, 
                null, 
                '🎨 Kontent tayyor! Mavzuga mos rasmlar yuklanmoqda va slayd dizayni shakllantirilmoqda...'
            );

            let pptx = new pptxgen();
            pptx.layout = 'LAYOUT_16x9';

            for (let i = 0; i < slidesData.length; i++) {
                const slideData = slidesData[i];
                let slide = pptx.addSlide();
                
                slide.background = { fill: "F8F9FA" };

                // Pastki qismdagi chiziq (Dizayn)
                slide.addShape(pptx.ShapeType.rect, {
                    x: 0, y: 7.0, w: "100%", h: 0.5,
                    fill: { color: themeColor }
                });

                // Sarlavha
                slide.addText(slideData.title, { 
                    x: 0.5, y: 0.5, w: "90%", h: 0.8, 
                    fontSize: 28, bold: true, color: themeColor,
                    fontFace: "Arial"
                });
                
                const hasChart = slideData.chart && slideData.chart.labels && slideData.chart.values && slideData.chart.labels.length > 0;
                
                // Matn qismi
                const bodyText = (slideData.bullets || []).join('\n\n');
                slide.addText(bodyText, { 
                    x: 0.5, y: 1.5, w: hasChart ? "45%" : "50%", h: 5.0, 
                    fontSize: hasChart ? 16 : 18, color: "334155", 
                    bullet: { type: 'number' },
                    fontFace: "Arial",
                    valign: 'top',
                    wrap: true
                });

                if (hasChart) {
                    // Grafika qo'shish
                    const chartData = [
                        {
                            name: slideData.chart.title || 'Statistika',
                            labels: slideData.chart.labels,
                            values: slideData.chart.values
                        }
                    ];
                    
                    const chartType = slideData.chart.type === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
                    
                    slide.addChart(chartType, chartData, {
                        x: 5.2, y: 1.5, w: 4.5, h: 4.5,
                        showLegend: true, legendPos: 'b',
                        showTitle: true, title: slideData.chart.title, titleFontSize: 14, titleColor: "334155",
                        barDir: 'col',
                        showValue: true,
                        chartColors: [themeColor, "3B82F6", "10B981", "F59E0B", "EF4444"]
                    });
                } else {
                    // Rasm qo'shish
                    const imgFilename = `temp_img_${Date.now()}_${i}.jpg`;
                    const searchKeyword = slideData.keyword || "education";
                    const defaultImgUrl = `https://images.unsplash.com/photo-1546410531-bb4caa6b424d?q=80&w=600&auto=format&fit=crop`;
                    const dynamicImageUrl = `https://loremflickr.com/600/450/${encodeURIComponent(searchKeyword)}`;

                    try {
                        await downloadImage(dynamicImageUrl, imgFilename);
                    } catch (imgErr) {
                        try {
                            await downloadImage(defaultImgUrl, imgFilename);
                        } catch(e) {}
                    }

                    if (fs.existsSync(imgFilename)) {
                        slide.addImage({
                            path: imgFilename,
                            x: 5.5, y: 1.5, w: 4.0, h: 4.5,
                            sizing: { type: 'cover', w: 4.0, h: 4.5 }
                        });
                        setTimeout(() => { try { fs.unlinkSync(imgFilename); } catch(e){} }, 5000);
                    }
                }
            }

            const fileName = path.join(__dirname, `${Date.now()}_taqdimot.pptx`);
            await pptx.writeFile({ fileName: fileName });

            await ctx.replyWithDocument({ 
                source: fileName, 
                filename: `${topic.replace(/\s+/g, '_')}_taqdimot.pptx` 
            });

            fs.unlinkSync(fileName);
            await ctx.telegram.deleteMessage(ctx.chat.id, loadingMessage.message_id);

        } catch (error) {
            console.error("Botda xatolik:", error);
            await ctx.telegram.editMessageText(
                ctx.chat.id, 
                loadingMessage.message_id, 
                null, 
                `❌ Xatolik yuz berdi: ${error.message}`
            );
        }
    });

    bot.launch().then(() => {
        console.log('🤖 Telegram Bot (backend ichida) ishga tushdi...');
    }).catch(err => {
        console.error("Botni ishga tushirishda xatolik:", err);
    });

    process.once('SIGINT', () => bot.stop('SIGINT'));
    process.once('SIGTERM', () => bot.stop('SIGTERM'));
}
