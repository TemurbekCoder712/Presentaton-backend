/**
 * ============================================================
 * PRESENTATION AI — 2-Bosqichli Pipeline
 * ============================================================
 * Bosqich 1: DeepSeek  → Slayd kontent + struktura (JSON)
 * Bosqich 2: Gemini    → Chiroyli HTML slaydlar (Reveal.js)
 * ============================================================
 * Ikkala model ham MegaLLM API orqali MEGALLM_API_KEY bilan chaqiriladi
 */

import { OpenAI } from 'openai';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.resolve(__dirname, '../.env') });

// ─── MegaLLM Client — Bosqich 1 (DeepSeek) ─────────────────
const deepseekClient = new OpenAI({
    baseURL: 'https://ai.megallm.io/v1',
    apiKey: process.env.MEGALLM_API_KEY,
});

// ─── MegaLLM Client — Bosqich 2 (Gemini yoki DeepSeek fallback)
const geminiClient = new OpenAI({
    baseURL: 'https://ai.megallm.io/v1',
    apiKey: process.env.MEGALLM_API_KEY,
});

// Model nomlarini shu yerda o'zgartiring:
const MODELS = {
    content: 'deepseek-ai/deepseek-v3.1',   // Kontent yaratish
    design: 'gemini-2.5-flash',              // HTML dizayn ✅ MegaLLM da bor
};

// ─── SYSTEM PROMPTS ────────────────────────────────────────

const DEEPSEEK_SYSTEM = `Siz professional taqdimotlar bo'yicha kontent-arxitektorsiz. Foydalanuvchi bergan mavzu bo'yicha 5 ta slayddan iborat mukammal taqdimot mazmunini tuzing. Har bir slayd uchun quyidagilarni kiriting:
- title: slayd sarlavhasi
- subtitle: qisqa ta'rif (ixtiyoriy)
- bullets: kamida 3 ta asosiy fikr (har biri 8-12 so'z)
- speaker_notes: spiker bu slayd haqida nima deyishi kerak (2-3 jumlа)
Vizual dizayn va kod YOZMANG. Faqat kontentning mantiqiy to'liqligi va boyligiga e'tibor bering.
JAVOB FAQAT JSON ARRAYDA BO'LSIN: [{"title":"...","subtitle":"...","bullets":["..."],"speaker_notes":"..."}]`;

const GEMINI_SYSTEM = `Siz UI/UX va Presentation Design ekspertisiz. Sizga berilgan JSON formatidagi taqdimot mazmunini qabul qiling va uni to'liq, self-contained HTML fayliga aylantiring. Quyidagi talablar majburiy:
1. Reveal.js (CDN orqali) ishlatilsin
2. Har bir JSON obyekti → bitta Reveal.js section (<section>) bo'lsin
3. Zamonaviy gradient fon (quyuq moviy-binafsha: #0f0c29 → #302b63 → #24243e)
4. Chiroyli kartochkalar, border-radius, box-shadow
5. Google Fonts (Inter yoki Poppins) CDN orqali
6. Bullet pointlar oldida rangli emoji/ikon
7. speaker_notes → <aside class="notes"> ichida bo'lsin
8. Responsive va fullscreen
FAQAT YAKUNIY HTML KODNI QAYTARING. Hech qanday izoh yoki markdown yozmang.`;

// ─── BOSQICH 1: DeepSeek → JSON Kontent ───────────────────

async function generateContent(userPrompt) {
    console.log('\n⚙️  [Bosqich 1] DeepSeek: kontent yaratilmoqda...');

    const response = await megaClient.chat.completions.create({
        model: MODELS.content,
        messages: [
            { role: 'system', content: DEEPSEEK_SYSTEM },
            { role: 'user', content: `Mavzu: "${userPrompt}"` }
        ],
        max_tokens: 1500,
        temperature: 0.7,
    });

    let raw = response.choices[0].message.content.trim();

    // Markdown ```json ... ``` ni tozalaymiz
    if (raw.startsWith('```')) {
        raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/, '').trim();
    }

    const slides = JSON.parse(raw);
    console.log(`✅ [Bosqich 1] ${slides.length} ta slayd kontent tayyor`);
    return slides;
}

// ─── BOSQICH 2: Gemini → HTML Dizayn ──────────────────────

async function generateDesign(slides, topic) {
    console.log(`\n🎨 [Bosqich 2] HTML dizayn yaratilmoqda (${MODELS.design})...`);

    const slidesJson = JSON.stringify(slides, null, 2);

    try {
        const response = await geminiClient.chat.completions.create({
            model: MODELS.design,
            messages: [
                { role: 'system', content: GEMINI_SYSTEM },
                {
                    role: 'user',
                    content: `Taqdimot mavzusi: "${topic}"\n\nSlayd kontent (JSON):\n${slidesJson}`
                }
            ],
            max_tokens: 4000,
            temperature: 0.4,
        });

        let html = response.choices[0].message.content.trim();

        // Markdown wrapper tozalaymiz
        if (html.startsWith('```')) {
            html = html.replace(/^```html\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/, '').trim();
        }

        if (html.includes('<html') || html.includes('<!DOCTYPE')) {
            console.log('✅ [Bosqich 2] HTML dizayn tayyor');
            return html;
        }

        console.warn('⚠️  Model HTML qaytarmadi, fallback...');
        return buildFallbackHtml(slides, topic);

    } catch (err) {
        console.warn(`⚠️  Dizayn modeli xato (${err.message}), fallback HTML..`);
        return buildFallbackHtml(slides, topic);
    }
}

// ─── FALLBACK HTML (Gemini ishlamasa) ─────────────────────

function buildFallbackHtml(slides, topic) {
    const sectionsHtml = slides.map((s, i) => `
        <section>
            <div class="slide-card">
                <div class="slide-num">0${i + 1}</div>
                <h2>${s.title || ''}</h2>
                ${s.subtitle ? `<p class="subtitle">${s.subtitle}</p>` : ''}
                <ul>
                    ${(s.bullets || []).map(b => `<li>${b}</li>`).join('')}
                </ul>
            </div>
            ${s.speaker_notes ? `<aside class="notes">${s.speaker_notes}</aside>` : ''}
        </section>
    `).join('');

    return `<!DOCTYPE html>
<html lang="uz">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${topic} — Presentation AI</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/5.1.0/reveal.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/5.1.0/theme/black.min.css">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  :root { --r-main-font: 'Inter', sans-serif; --r-heading-font: 'Inter', sans-serif; }
  .reveal { background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); }
  .slide-card { background: rgba(255,255,255,0.07); border-radius: 20px; padding: 40px; backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.12); }
  .slide-num { font-size: 13px; font-weight: 700; color: #a78bfa; letter-spacing: 0.15em; margin-bottom: 14px; text-transform: uppercase; }
  .reveal h2 { font-size: 1.8em; font-weight: 800; color: #fff; margin: 0 0 12px; line-height: 1.2; }
  .subtitle { color: #c4b5fd; font-size: 1em; margin: 0 0 24px; }
  .reveal ul { list-style: none; padding: 0; margin: 0; text-align: left; }
  .reveal ul li { padding: 8px 0 8px 32px; position: relative; color: #e0e7ff; font-size: 0.85em; line-height: 1.6; border-bottom: 1px solid rgba(255,255,255,0.06); }
  .reveal ul li::before { content: "▶"; position: absolute; left: 0; color: #f59e0b; font-size: 12px; top: 10px; }
</style>
</head>
<body>
<div class="reveal">
  <div class="slides">
    <section>
      <div class="slide-card">
        <div class="slide-num">Presentation AI</div>
        <h2>${topic}</h2>
        <p class="subtitle">AI tomonidan tayyorlangan professional taqdimot</p>
      </div>
    </section>
    ${sectionsHtml}
  </div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/5.1.0/reveal.min.js"></script>
<script>
  Reveal.initialize({ hash: true, transition: 'slide', backgroundTransition: 'fade', controls: true, progress: true, center: true });
</script>
</body>
</html>`;
}

// ─── ASOSIY FUNKSIYA: generatePresentation ─────────────────

/**
 * 2-bosqichli pipeline
 * @param {string} userPrompt  — Foydalanuvchi yozgan mavzu/prompt
 * @returns {{ slides: Array, html: string }}
 */
export async function generatePresentation(userPrompt) {
    // Bosqich 1: DeepSeek → JSON kontent
    const slides = await generateContent(userPrompt);

    // Bosqich 2: Gemini → HTML dizayn
    const html = await generateDesign(slides, userPrompt);

    return { slides, html };
}
