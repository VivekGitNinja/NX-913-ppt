/* NX-913 Pitch Deck → PPTX generator
 * Parses kicker/title/subtitle/notes from index.html so headers stay in sync,
 * lays out curated body content with the deck's red-on-black branding. */
import fs from 'node:fs';
import path from 'node:path';
import pptxgen from 'pptxgenjs';

const ROOT = path.resolve(import.meta.dirname, '../..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

/* ---------- palette (matches deck + nx-913.com) ---------- */
const C = {
  bg: '050505', panel: '0F0F14', panel2: '19191E', border: '26262D',
  text: 'F5F5F5', muted: '9CA3AF', dim: '71717A',
  red: 'DC2626', red2: 'EF4444', red3: 'F87171', gold: 'FFD700', green: '22C55E'
};
const FH = 'Chakra Petch', FB = 'Rajdhani', FM = 'JetBrains Mono';

/* ---------- tiny html helpers ---------- */
const clean = s => s
  .replace(/<[^>]+>/g, '')
  .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, ' ')
  .replace(/\s+/g, ' ').trim();

/* ---------- parse slides from the live deck ---------- */
const sections = [...html.matchAll(/<section class="slide"([\s\S]*?)<\/section>/g)].map(m => m[1]);
const deckSlides = sections.map(sec => ({
  title: /data-title="([^"]*)"/.exec(sec)?.[1] || '',
  notes: /data-notes="([^"]*)"/.exec(sec)?.[1] || '',
  kicker: clean(/<div class="kicker[^"]*"[^>]*>([\s\S]*?)<\/div>/.exec(sec)?.[1] || ''),
  sTitle: clean(/<h2 class="s-title[^"]*"[^>]*>([\s\S]*?)<\/h2>/.exec(sec)?.[1] || ''),
  sSub: clean(/<p class="s-sub[^"]*"[^>]*>([\s\S]*?)<\/p>/.exec(sec)?.[1] || '')
}));

/* ---------- curated body content (from the master blueprint) ---------- */
const BODIES = [
  // 1 — Title & Vision (special layout)
  null,
  // 2 — Campus Dilemma
  [
    { h: '🧩 Scattered Information', p: '20+ clubs use separate Google Forms, WhatsApp groups, Notion pages & Unstop links — students miss out on events.' },
    { h: '💸 Manual Payment Chaos', p: 'UPI screenshots & manual bank verification → reconciliation delays, cash leakage & accounting disputes.' },
    { h: '🚪 Gate Congestion & Fake Passes', p: 'Manual paper registers at auditorium doors cause long queues & ticket duplication.' },
    { h: '⚔️ Hackathon Nightmares', p: 'Team formation friction, lost submissions & manual judging sheets → delayed results & disputes.' },
    { h: '📜 Certificate & Reporting Burden', p: 'Organizers spend weeks emailing Canva certificates; faculty spend weeks compiling NAAC/IQAC reports.' }
  ],
  // 3 — System Overview
  [
    { h: '🎓 Student / Attendee Portal', p: 'One-click event discovery, team builder, digital ticket pass & verifiable certificate vault.' },
    { h: '🎪 Organizer & Faculty Coordinator Portal', p: 'Event builder, ticketing tiers, coupon engine, broadcast center & live attendance scanner.' },
    { h: '⚖️ Hackathon Judging & Mentor Console', p: 'Multi-criteria weighted scoring rubrics, submission inspection & live dynamic leaderboards.' },
    { h: '🏛️ College Admin & Dean Oversight Panel', p: 'Event approval workflows, cross-department analytics, asset libraries & system audit logs.' }
  ],
  // 4 — Student Experience
  [
    { h: '🔐 Secure & Effortless Auth', p: 'Google OAuth, GitHub OAuth, OTP login & 2FA with session management.' },
    { h: '🎫 Ticket & Pass Wallet', p: 'Instant PDF tickets with QR codes, add-to-calendar (.ics) & live pass status.' },
    { h: '🤝 Solo Hacker & Team Matchmaking', p: 'Find teammates by skill tag — Frontend, AI/ML, Backend, UI/UX.' },
    { h: '🏆 Gamification & Ranks', p: 'Experience badges, departmental ranks & verified activity logs.' },
    { h: '🌐 Social & Community', p: 'Follow clubs, bookmark events, verified testimonials & upvote showcase projects.' },
    { h: '👤 Public Profile · /u/username', p: 'On-platform portfolio of badges, ranks & verified achievements.' }
  ],
  // 5 — Hackathon Engine
  [
    { h: '👥 Dynamic Team Workspace', p: 'Create teams, share auto-generated secure invite codes & assign project roles.' },
    { h: '📦 Submission Engine', p: 'Submit GitHub repos, live demo URLs, pitch videos, tech-stack tags & docs.' },
    { h: '⚖️ Digital Judging Portal', p: 'Multi-track judge assignment & weighted rubrics — Innovation 30% · Feasibility 30% · Code Quality 20% · Presentation 20%.' },
    { h: '🥇 Real-Time Dynamic Leaderboard', p: 'Auto-computed rankings, tie-breaker handling & instant podium reveals — zero grading disputes.' }
  ],
  // 6 — Smart Gate Check-in
  [
    { h: '📷 In-Browser QR Scanner', p: 'Works on any smartphone camera or webcam with sub-second verification — no special hardware.' },
    { h: '⏱️ Multi-Session Attendance', p: 'Track checkpoints across Day 1 Inauguration, Lunch, Day 2 Hackathon & Valedictory.' },
    { h: '📶 PWA & Offline Fallback Protocol', p: 'Campus Wi-Fi fails? Scanner caches the guest list in encrypted local storage (IndexedDB), validates offline, auto-syncs on reconnect.' },
    { h: '⚡ Zero Gate Delays', p: 'Sub-second verification, zero duplicate check-ins — gates never stop moving.' }
  ],
  // 7 — Certificate Studio
  [
    { h: '🎨 Visual Certificate Studio', p: 'Custom templates for Participants, Winners, Runners-Up, Mentors & Volunteers with {{name}}, {{event}}, {{date}}, {{rank}} placeholders.' },
    { h: '⚡ Bulk Queue Dispatch', p: 'BullMQ + Puppeteer workers generate hundreds of high-res PDF/PNG certificates in the background.' },
    { h: '🔍 Tamper-Proof QR Verification', p: 'Every certificate carries a unique cryptographic QR linking to the official /verify page — zero forgery.' },
    { h: '💼 Placement & LinkedIn Ready', p: 'One-click sharing to LinkedIn to boost student employability & showcase college achievements.' }
  ],
  // 8 — AI NAAC/IQAC Report Builder
  [
    { h: '📄 One-Click Official Export', p: 'Standardized institutional report templates in DOCX & PDF.' },
    { h: '📊 Pre-Populated Sections', p: 'Executive summary, participant lists with timestamps, speaker profiles, photo collage, feedback & financial balance sheet.' },
    { h: '🤖 AI-Assisted Drafting', p: 'Groq LPU + Gemini compile executive summaries & outcome analysis from raw event data.' },
    { h: '🏛️ Accreditation Compliant', p: 'Direct adherence to NAAC Criterion V, IQAC & AICTE/NIRF documentation guidelines.' }
  ],
  // 9 — Payments
  [
    { h: '💳 Integrated Razorpay Gateway', p: 'UPI (GPay, PhonePe, Paytm), Credit/Debit Cards & Net Banking.' },
    { h: '🎟️ Flexible Ticketing & Coupons', p: 'Free tickets, multi-tier pricing (Early Bird · Standard · VIP) & discount coupon management.' },
    { h: '🧾 Automated GST & Invoicing', p: 'Instant digital invoices & payment receipts issued to attendees.' },
    { h: '🛡️ Zero Cash Leakage & Audit Logs', p: 'Tamper-proof transaction logs for faculty & accounts, with automated bulk refunds.' }
  ],
  // 10 — Organizer Suite
  [
    { h: '📢 Email Broadcast Engine', p: 'Segmented 1-click updates to registered attendees, shortlisted teams or everyone.' },
    { h: '📁 Platform Asset Library', p: 'Central repository for college banners, department logos, sponsor decks & guidelines.' },
    { h: '📈 Live Analytics & Intelligence', p: 'Registration velocity, drop-off rates, revenue & venue capacity in real time.' },
    { h: '🤖 24/7 AI Attendee Assistant', p: 'Low-latency Groq + Gemini chatbot answering schedule, venue, rules & contact FAQs.' }
  ],
  // 11 — Architecture & Security
  [
    { h: '⚡ Frontend', p: 'Next.js 16 (App Router, SSR/SSG) · React 19 · Tailwind CSS · Zustand · Framer Motion · PWA offline.' },
    { h: '🛠️ Backend', p: 'Node.js 22 · Express.js · TypeScript · Mongoose ODM · REST API.' },
    { h: '🔄 Async Processing', p: 'Redis & BullMQ workers — Puppeteer PDF rendering, certificate generation, Nodemailer/Resend bulk mailers.' },
    { h: '🔒 Enterprise Security', p: 'JWT + refresh rotation, HTTP-only cookies, RBAC & CSRF, rate limiting, Helmet, Zod validation, audit logs & Sentry.' }
  ],
  // 12 — Admin Governance
  [
    { h: '🛡️ Event Approval Workflow', p: 'Club events require faculty/Dean review before publishing — no unauthorized listings.' },
    { h: '🏢 Club & Organizer Verification', p: 'Verify clubs, assign organizer permissions & audit coordinator actions.' },
    { h: '📊 Cross-Department Insights', p: 'Compare participation across CSE, IT, ECE, Mech & MBA; annual activity summaries.' },
    { h: '🎛️ Platform Feature Flags', p: 'Enable/disable payments, hackathons & public signups dynamically via backend toggles.' }
  ],
  // 13 — Live Demo (special: 4 steps)
  [
    { h: '1️⃣ Student Flow', p: 'Browse events → Register in 1 click → QR Pass appears in wallet.' },
    { h: '2️⃣ Hackathon Flow', p: 'Form team with invite code → Submit demo URL & GitHub repo.' },
    { h: '3️⃣ Gate Scanner', p: 'Organizer mobile camera scans student QR → sub-second attendance confirmation.' },
    { h: '4️⃣ Admin / Faculty Flow', p: 'One-click certificate release with QR verification → one-click NAAC/IQAC report download.' }
  ],
  // 14 — Rollout Plan (special: phases + CTA)
  [
    { h: '📅 Phase 1 — Pilot Deployment', p: 'Weeks 1–2 · Deploy for 2 departmental workshops; validate gate scanning & certificate dispatch.' },
    { h: '📅 Phase 2 — Club & Faculty Onboarding', p: 'Weeks 3–4 · 30-min training for club leads, societies (IEEE, GDSC, ACM, Cultural) & coordinators.' },
    { h: '📅 Phase 3 — Campus-Wide Launch', p: 'Month 2 · Official event portal linked on the college website & student app.' },
    { h: '🎯 Our Request Today', p: 'Seeking official approval from College Administration to launch the Phase 1 Pilot.' }
  ],
  // 15 — Q&A (special: Q/A pairs)
  [
    { h: '"What if the internet fails during a 1,000-person fest?"', p: 'Offline PWA Protocol — encrypted local cache, sub-0.5s validation, auto-sync on reconnect.' },
    { h: '"Can a student create a fake or inappropriate event?"', p: 'Strict RBAC — only verified club organizers can create, and events need faculty/admin approval.' },
    { h: '"How does this help NAAC / NBA accreditation?"', p: 'AI report generator exports structured DOCX/PDF formatted for NAAC Criterion V.' },
    { h: '"Is student data secure and private?"', p: 'JWT with refresh rotation, 2FA, encrypted hashing, role-based queries & rate-limiting.' },
    { h: '"Will the server crash if 2,000 students register at once?"', p: 'Async BullMQ/Redis workers handle PDFs, certificates & mail — API stays light.' }
  ],
  // 16 — Thank You (special layout)
  null
];

/* ---------- pptx ---------- */
const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'NX-913';
pptx.title = 'NX-913 — Campus Event & Hackathon Operating System';
pptx.subject = 'Pitch Deck';

const W = 13.333, H = 7.5;

function brandTop(slide) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.45, y: 0.38, w: 0.42, h: 0.42, rectRadius: 0.08,
    fill: { color: C.red }, line: { type: 'none' }
  });
  slide.addText('N', { x: 0.45, y: 0.38, w: 0.42, h: 0.42, align: 'center', valign: 'middle',
    fontFace: FH, fontSize: 16, bold: true, color: 'FFFFFF' });
  slide.addText([{ text: 'NX-913  ', options: { fontFace: FH, bold: true, color: C.text } },
    { text: 'CAMPUS OS', options: { fontFace: FB, color: C.dim, fontSize: 11 } }],
    { x: 0.97, y: 0.4, w: 3.2, h: 0.38, valign: 'middle' });
  slide.addText('PITCH DECK', { x: W - 2.6, y: 0.44, w: 2.15, h: 0.3, align: 'right',
    fontFace: FM, fontSize: 9, color: C.dim, charSpacing: 2 });
}

function footer(slide, idx, total) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: H - 0.09, w: W * ((idx + 1) / total), h: 0.05, fill: { color: C.red }
  });
  slide.addText('NX-913 · CAMPUS EVENT & HACKATHON OS', { x: 0.45, y: H - 0.55, w: 6, h: 0.3,
    fontFace: FB, fontSize: 9, color: C.dim, charSpacing: 1.5 });
  slide.addText(String(idx + 1).padStart(2, '0') + ' / ' + String(total).padStart(2, '0'),
    { x: W - 1.6, y: H - 0.55, w: 1.15, h: 0.3, align: 'right', fontFace: FM, fontSize: 10, color: C.muted });
}

function header(slide, d, idx) {
  const kicker = d.kicker || ('SLIDE ' + String(idx + 1).padStart(2, '0'));
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.45, y: 1.05, w: 1.25, h: 0.38, rectRadius: 0.19,
    fill: { color: '1A0D0D' }, line: { color: '7F1D1D', width: 1 }
  });
  slide.addText(kicker.toUpperCase(), { x: 0.45, y: 1.05, w: 1.25, h: 0.38, align: 'center', valign: 'middle',
    fontFace: FM, fontSize: 9, color: C.red3, charSpacing: 1 });
  slide.addText(d.sTitle || d.title, { x: 0.42, y: 1.5, w: W - 0.9, h: 0.9, valign: 'middle',
    fontFace: FH, fontSize: 34, bold: true, color: C.text });
  if (d.sSub) {
    slide.addText(d.sSub, { x: 0.45, y: 2.32, w: W - 1.8, h: 0.45, valign: 'top',
      fontFace: FB, fontSize: 15, color: C.muted });
  }
}

function card(slide, x, y, w, h, item, accent, compact) {
  const hs = compact ? 12.5 : 14.5, ps = compact ? 11 : 12.5;
  const hy = compact ? 0.07 : 0.1, py = compact ? 0.44 : 0.58;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: C.panel }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.ShapeType.rect, { x, y: y + 0.12, w: 0.05, h: h - 0.24, fill: { color: accent || C.red } });
  slide.addText(item.h, { x: x + 0.22, y: y + hy, w: w - 0.4, h: 0.42, valign: 'top',
    fontFace: FH, fontSize: hs, bold: true, color: C.text, breakLine: false });
  if (item.p) {
    slide.addText(item.p, { x: x + 0.22, y: y + py, w: w - 0.4, h: h - py - 0.12, valign: 'top',
      fontFace: FB, fontSize: ps, color: C.muted, lineSpacing: 18 });
  }
}

function bodyGrid(slide, items, accent) {
  const top = 2.95, bottom = H - 0.75;
  const n = items.length;
  const cols = n === 5 ? 2 : 3;
  const rows = Math.ceil(n / cols);
  const gap = 0.22;
  const colW = (W - 0.9 - gap * (cols - 1)) / cols;
  const rowH = (bottom - top - gap * (rows - 1)) / rows;
  items.forEach((it, i) => {
    const r = Math.floor(i / cols), c = i % cols;
    card(slide, 0.45 + c * (colW + gap), top + r * (rowH + gap), colW, rowH, it, accent);
  });
}

/* ---------- framed screenshot (browser-chrome mockup) ---------- */
const IMG = {
  home: 'assets/site/01-home.png',
  events: 'assets/site/02-events.png',
  hackathons: 'assets/site/03-hackathons.png',
  detail: 'assets/site/04-hackathon-detail.png'
};

function framedShot(slide, img, x, y, w, h, url) {
  // outer frame
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.09,
    fill: { color: '0C0C10' }, line: { color: C.border2, width: 1 } });
  // chrome bar
  slide.addShape(pptx.ShapeType.rect, { x: x + 0.06, y: y + 0.06, w: w - 0.12, h: 0.34,
    fill: { color: '14141A' }, line: { type: 'none' } });
  ['7F1D1D', 'B45309', '166534'].forEach((c, i) => {
    slide.addShape(pptx.ShapeType.ellipse, { x: x + 0.14 + i * 0.2, y: y + 0.155, w: 0.12, h: 0.12,
      fill: { color: c }, line: { type: 'none' } });
  });
  slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.72, y: y + 0.13, w: w - 1.0, h: 0.2, rectRadius: 0.1,
    fill: { color: '0A0A0E' }, line: { color: C.border, width: 0.5 } });
  slide.addText(url || 'nx-913.com', { x: x + 0.78, y: y + 0.13, w: w - 1.1, h: 0.2, align: 'center', valign: 'middle',
    fontFace: FM, fontSize: 8, color: C.dim });
  // image
  slide.addImage({ path: path.join(ROOT, img), x: x + 0.06, y: y + 0.46, w: w - 0.12, h: h - 0.56 });
  // subtle glow under frame
  slide.addShape(pptx.ShapeType.ellipse, { x: x + w / 2 - 1.6, y: y + h - 0.35, w: 3.2, h: 0.5,
    fill: { color: C.red, transparency: 86 }, line: { type: 'none' } });
}

/* ---------- special layouts ---------- */
function titleSlide(slide, d, idx, total) {
  brandTop(slide);
  slide.addShape(pptx.ShapeType.ellipse, { x: -1.5, y: -1.2, w: 7, h: 7,
    fill: { color: C.red, transparency: 88 }, line: { type: 'none' } });
  // left: wordmark + copy
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.85, y: 2.0, w: 2.2, h: 0.46, rectRadius: 0.23,
    fill: { color: '120B0B' }, line: { color: '7F1D1D', width: 1 }
  });
  slide.addText('N X - 9 1 3', { x: 0.85, y: 2.0, w: 2.2, h: 0.46, align: 'center', valign: 'middle',
    fontFace: FM, fontSize: 11, color: C.red3, charSpacing: 1.5 });
  slide.addText('NX-913', { x: 0.82, y: 2.55, w: 6.6, h: 1.4, align: 'left', valign: 'middle',
    fontFace: FH, fontSize: 66, bold: true, color: C.red2, charSpacing: 1 });
  slide.addText('The Next-Generation Campus Event & Hackathon Operating System',
    { x: 0.85, y: 3.95, w: 7.1, h: 0.9, align: 'left', fontFace: FH, fontSize: 20, bold: true, color: C.text, lineSpacing: 28 });
  slide.addText('From Registration to NAAC Reports — In One Click',
    { x: 0.85, y: 5.0, w: 7.1, h: 0.4, align: 'left', fontFace: FB, fontSize: 15, color: C.gold });
  slide.addText('Unified Student Engagement · Hackathon Lifecycle · Academic Compliance',
    { x: 0.85, y: 5.5, w: 7.1, h: 0.4, align: 'left', fontFace: FB, fontSize: 13, color: C.muted });
  // chips
  const chips = ['🎓 Student Portal', '🎪 Organizer Suite', '⚖️ Judging Console', '🏛️ Admin Oversight'];
  chips.forEach((c, i) => {
    const cw = 1.62, gap = 0.18;
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.85 + i * (cw + gap), y: 6.05, w: cw, h: 0.42, rectRadius: 0.21,
      fill: { color: C.panel2 }, line: { color: C.border, width: 1 } });
    slide.addText(c, { x: 0.85 + i * (cw + gap), y: 6.05, w: cw, h: 0.42, align: 'center', valign: 'middle',
      fontFace: FB, fontSize: 10.5, bold: true, color: C.text });
  });
  // right: live screenshot of the real product
  framedShot(slide, IMG.home, 8.35, 1.35, 4.4, 5.35, 'https://nx-913.com');
  footer(slide, idx, total);
}

function sideShotSlide(slide, d, idx, total, items, img, url, accent) {
  brandTop(slide);
  header(slide, d, idx);
  // cards on the left
  const leftW = 7.7;
  const top = 3.0, bottom = H - 0.75;
  const n = items.length;
  const cols = 2;
  const rows = Math.ceil(n / cols);
  const gap = 0.2;
  const colW = (leftW - gap) / 2;
  const rowH = (bottom - top - gap * (rows - 1)) / rows;
  const compact = rowH < 1.35;
  items.forEach((it, i) => {
    const r = Math.floor(i / cols), c = i % cols;
    card(slide, 0.45 + c * (colW + gap), top + r * (rowH + gap), colW, rowH, it, accent, compact);
  });
  // screenshot on the right
  framedShot(slide, img, 8.4, 2.75, 4.35, 4.15, url);
}


function thanksSlide(slide, d, idx, total) {
  brandTop(slide);
  slide.addShape(pptx.ShapeType.ellipse, { x: W - 5.5, y: -1.5, w: 7, h: 7,
    fill: { color: C.red, transparency: 90 }, line: { type: 'none' } });
  slide.addText('Thank You', { x: 0, y: 2.2, w: W, h: 1.5, align: 'center', valign: 'middle',
    fontFace: FH, fontSize: 66, bold: true, color: C.text });
  slide.addText('NX-913 — the unified operating system for every campus event, hackathon & accreditation report.',
    { x: W / 2 - 4.2, y: 3.95, w: 8.4, h: 0.55, align: 'center', fontFace: FB, fontSize: 16, color: C.muted });
  slide.addText('“From Registration to NAAC Reports — In One Click.”',
    { x: W / 2 - 3.5, y: 4.6, w: 7, h: 0.4, align: 'center', fontFace: FB, fontSize: 13.5, color: C.gold });
  const chips = ['🎟️ Events & Passes', '🏆 Hackathons', '📜 Certificates', '📄 NAAC Reports'];
  chips.forEach((c, i) => {
    const cw = 2.35, gap = 0.3, totalW = chips.length * cw + (chips.length - 1) * gap;
    const x = W / 2 - totalW / 2 + i * (cw + gap);
    slide.addShape(pptx.ShapeType.roundRect, { x, y: 5.3, w: cw, h: 0.5, rectRadius: 0.25,
      fill: { color: C.panel2 }, line: { color: C.border, width: 1 } });
    slide.addText(c, { x, y: 5.3, w: cw, h: 0.5, align: 'center', valign: 'middle',
      fontFace: FB, fontSize: 12, bold: true, color: C.text });
  });
  slide.addText('QUESTIONS & DISCUSSION — OPEN FLOOR', { x: W / 2 - 3, y: 6.15, w: 6, h: 0.4, align: 'center',
    fontFace: FM, fontSize: 11, color: C.red3, charSpacing: 2 });
  footer(slide, idx, total);
}

/* ---------- build ---------- */
const total = deckSlides.length;
deckSlides.forEach((d, i) => {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  const body = BODIES[i];

  if (i === 0) { titleSlide(slide, d, i, total); }
  else if (i === total - 1) { thanksSlide(slide, d, i, total); }
  else if (i === 3) { sideShotSlide(slide, d, i, total, body, IMG.events, 'https://nx-913.com/events'); }
  else if (i === 4) { sideShotSlide(slide, d, i, total, body, IMG.detail, 'https://nx-913.com/hackathons'); }
  else if (i === 12) { sideShotSlide(slide, d, i, total, body, IMG.hackathons, 'https://nx-913.com/hackathons'); }
  else {
    brandTop(slide);
    header(slide, d, i);
    if (body) bodyGrid(slide, body);
  }
  if (d.notes) slide.addNotes(d.notes);
});

const out = path.join(ROOT, 'NX-913_Pitch_Deck.pptx');
await pptx.writeFile({ fileName: out });
console.log('Wrote', out, fs.statSync(out).size, 'bytes');
