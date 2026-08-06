import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Google Sheets ID
const SPREADSHEET_ID = '1ypPq4dyGLQne_w00rTbIEAqqJ7u7_TDLC8KCuK8LVxw';

// Credentials file
const credsPath = path.join(__dirname, 'credentials.json');
let creds = null;
if (fs.existsSync(credsPath)) {
    creds = JSON.parse(fs.readFileSync(credsPath, 'utf8'));
}

async function appendUserToSheet(user) {
    if (!creds) {
        console.log("⚠️ Credentials fayli yo'q, Google Sheets ga yozilmadi.");
        return;
    }

    try {
        const serviceAccountAuth = new JWT({
            email: creds.client_email,
            key: creds.private_key,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo(); 
        const sheet = doc.sheetsByIndex[0]; 
        
        const rows = await sheet.getRows();
        const targetId = user.id?.toString();
        
        let existingRow = null;
        for (let i = rows.length - 1; i >= 0; i--) {
            if (rows[i].get('Telegram ID') === targetId) {
                existingRow = rows[i];
                break;
            }
        }

        if (existingRow) {
            // Agar foydalanuvchi allaqachon jadvalda bo'lsa, faqat eng oxirgi kirgan sanasini yangilaymiz
            // Lekin holatini yoki boshqa narsalarini o'zgartirmaymiz (dublikat bo'lmaydi)
            existingRow.set('Sana', new Date().toLocaleString('ru-RU'));
            await existingRow.save();
            console.log(`✅ Foydalanuvchi jadvalda bor, faqat sanasi yangilandi: ${user.first_name}`);
        } else {
            // Yangi foydalanuvchini qo'shamiz
            await sheet.addRow({
                'Telegram ID': targetId || '',
                'Ism Familiya': `${user.first_name || ''} ${user.last_name || ''}`.trim() || 'No Name',
                'Username': user.username ? `@${user.username}` : 'No username',
                'Sana': new Date().toLocaleString('ru-RU'),
                'Holati': 'Faqat /start bosgan'
            });
            console.log(`✅ Yangi foydalanuvchi Google Sheets'ga yozildi: ${user.first_name}`);
        }
    } catch (error) {
        console.error('❌ Google Sheets ga yozishda xatolik:', error.message);
    }
}

async function appendFeedbackToSheet(name, username, feedback) {
    if (!creds) {
        console.log("⚠️ Credentials fayli yo'q, Google Sheets ga yozilmadi.");
        return;
    }

    try {
        const serviceAccountAuth = new JWT({
            email: creds.client_email,
            key: creds.private_key,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo(); 
        
        // 2-sahifa (Лист2 yoki Feedbacklar) index 1 da joylashgan
        if (doc.sheetCount < 2) {
            console.error("❌ Google Sheets da 2-sahifa (Feedbacklar) topilmadi!");
            return;
        }
        
        const sheet = doc.sheetsByIndex[1]; 
        
        // Add row
        await sheet.addRow({
            'Ism': name || 'Noma\'lum',
            'Username': username ? (username.startsWith('@') ? username : `@${username}`) : 'No username',
            'Fikr': feedback || '',
            'Sana': new Date().toLocaleString('ru-RU')
        });
        
        console.log(`✅ Feedback Google Sheets'ga yozildi: ${name}`);
    } catch (error) {
        console.error('❌ Feedbackni Google Sheets ga yozishda xatolik:', error.message);
    }
}

async function updateUserStatus(telegramId, topic = null) {
    if (!creds || !telegramId) return;
    try {
        const serviceAccountAuth = new JWT({
            email: creds.client_email,
            key: creds.private_key,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo(); 
        const sheet = doc.sheetsByIndex[0];
        
        const rows = await sheet.getRows();
        const targetId = telegramId.toString();
        
        for (let i = rows.length - 1; i >= 0; i--) {
            if (rows[i].get('Telegram ID') === targetId) {
                let updated = false;
                
                // 1. Holatni yangilash
                if (rows[i].get('Holati') !== '/start bosgan va ilovani ishlatgan') {
                    rows[i].set('Holati', '/start bosgan va ilovani ishlatgan');
                    updated = true;
                }
                
                // 2. Mavzuni qo'shish
                if (topic) {
                    try {
                        const oldTopic = rows[i].get('Taqdimot mavzusi') || '';
                        // Agar bu mavzu avval qo'shilmagan bo'lsa
                        if (!oldTopic.includes(topic)) {
                            const newTopic = oldTopic ? `${oldTopic}, ${topic}` : topic;
                            rows[i].set('Taqdimot mavzusi', newTopic);
                            updated = true;
                        }
                    } catch (headerErr) {
                        // Agar "Taqdimot mavzusi" degan header jadvalda bo'lmasa xato beradi, shuni ushlaymiz
                        console.log("Jadvalda 'Taqdimot mavzusi' ustuni topilmadi.");
                    }
                }

                if (updated) {
                    await rows[i].save();
                    console.log(`✅ Foydalanuvchi holati yangilandi: ${targetId}`);
                }
                break;
            }
        }
    } catch (error) {
        console.error("Holatni yangilashda xatolik:", error.message);
    }
}

export { appendUserToSheet, appendFeedbackToSheet, updateUserStatus };
