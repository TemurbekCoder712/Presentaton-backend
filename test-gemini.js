import { GoogleGenerativeAI } from '@google/generative-ai';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, '../.env') });

async function test() {
    try {
        console.log("Gemini kaliti:", !!process.env.GEMINI_API_KEY);
        const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
        const geminiModel = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
        console.log("Model oldi, so'rov yuborilmoqda...");
        const result = await geminiModel.generateContent("Salom, qandaysan?");
        console.log("Javob keldi:", result.response.text());
    } catch (e) {
        console.error("Xatolik yuz berdi:", e);
    }
}
test();
