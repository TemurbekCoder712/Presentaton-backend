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
        
        const targetId = user.id?.toString();
        
        await sheet.addRow({
            'Telegram ID': targetId || '',
            'Ism Familiya': `${user.first_name || ''} ${user.last_name || ''}`.trim() || 'No Name',
            'Username': user.username ? (user.username.startsWith('@') ? user.username : `@${user.username}`) : 'No username',
            'Sana': new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' }),
            'Holati': 'Faqat /start bosgan'
        });
        console.log(`✅ Yangi foydalanuvchi (/start) Google Sheets'ga yozildi: ${user.first_name}`);
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
            'Sana': new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' })
        });
        
        console.log(`✅ Feedback Google Sheets'ga yozildi: ${name}`);
    } catch (error) {
        console.error('❌ Feedbackni Google Sheets ga yozishda xatolik:', error.message);
    }
}

async function appendQuestionToSheet(telegramId, name, username, question, aiResponse) {
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
        
        // 3-sahifa (Лист3 yoki Savollar) index 2 da joylashgan
        if (doc.sheetCount < 3) {
            console.error("❌ Google Sheets da 3-sahifa (Savollar) topilmadi!");
            return;
        }
        
        const sheet = doc.sheetsByIndex[2]; 
        
        // Add row
        await sheet.addRow({
            'Telegram ID': telegramId?.toString() || '',
            'Ism Familiya': name || 'Noma\'lum',
            'Username': username ? (username.startsWith('@') ? username : `@${username}`) : 'No username',
            'Sana': new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' }),
            'Bergan savoli': question || '',
            'Support Javobi': aiResponse || ''
        });
        
        console.log(`✅ Savol Google Sheets'ga yozildi: ${name}`);
    } catch (error) {
        console.error('❌ Savolni Google Sheets ga yozishda xatolik:', error.message);
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
        
        let lastKnownName = 'Noma\'lum';
        let lastKnownUsername = 'No username';

        // Oxirgi ma'lumotlarini qidiramiz
        for (let i = rows.length - 1; i >= 0; i--) {
            if (rows[i].get('Telegram ID') === targetId) {
                lastKnownName = rows[i].get('Ism Familiya') || lastKnownName;
                lastKnownUsername = rows[i].get('Username') || lastKnownUsername;
                break;
            }
        }

        await sheet.addRow({
            'Telegram ID': targetId,
            'Ism Familiya': lastKnownName,
            'Username': lastKnownUsername,
            'Sana': new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' }),
            'Holati': topic ? 'Taqdimot yaratdi' : 'Ilovaga kirdi',
            'Taqdimot mavzusi': topic || ''
        });
        
        console.log(`✅ Foydalanuvchi harakati qo'shildi: ${targetId} (${topic ? 'Taqdimot' : 'Tashrif'})`);
    } catch (error) {
        console.error("Holatni yangilashda xatolik:", error.message);
    }
}

export { appendUserToSheet, appendFeedbackToSheet, updateUserStatus, appendQuestionToSheet };
