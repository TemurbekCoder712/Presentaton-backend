import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Google Sheets ID
const SPREADSHEET_ID = '1ypPq4dyGIQne_wC0rTbIFAqqj7u7_TDiC8KCuK8IVxw';

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
            'Sana': new Date().toLocaleString('ru-RU')
        });
        
        console.log(`✅ Foydalanuvchi Google Sheets'ga yozildi: ${user.first_name}`);
    } catch (error) {
        console.error('❌ Google Sheets ga yozishda xatolik:', error.message);
    }
}

export { appendUserToSheet };
