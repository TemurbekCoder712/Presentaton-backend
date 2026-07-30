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
        
        // Add row
        await sheet.addRow({
            'Telegram ID': user.id?.toString() || '',
            'Ism Familiya': `${user.first_name || ''} ${user.last_name || ''}`.trim() || 'No Name',
            'Username': user.username ? `@${user.username}` : 'No username',
            'Sana': new Date().toLocaleString('ru-RU'),
            'Holati': 'Faqat /start bosgan'
        });
        
        console.log(`✅ Foydalanuvchi Google Sheets'ga yozildi: ${user.first_name}`);
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

async function updateUserStatus(telegramId) {
    if (!creds) return;
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
                if (rows[i].get('Holati') !== 'Ilovani ishlatgan') {
                    rows[i].set('Holati', 'Ilovani ishlatgan');
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
