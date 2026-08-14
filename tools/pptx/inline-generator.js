/* ============ IN-DECK DOWNLOAD PPT (client-side pptxgenjs) ============ */
(function(){
  if (window.__nxPPTInstalled) return;
  window.__nxPPTInstalled = true;

  /* base64 site screenshots (compressed, injected at build time) */
  var SHOTS = {
    home:  '__B64_HOME__',
    events: '__B64_EVENTS__',
    hackathons: '__B64_HACKATHONS__',
    detail: '__B64_DETAIL__'
  };

  /* curated body content, mirrors tools/pptx/export-pptx.mjs */
  var BODIES = [
    null,
    [
      { h: '🧩 Scattered Information', p: '20+ clubs use separate Google Forms, WhatsApp groups, Notion pages & Unstop links — students miss out on events.' },
      { h: '💸 Manual Payment Chaos', p: 'UPI screenshots & manual bank verification → reconciliation delays, cash leakage & accounting disputes.' },
      { h: '🚪 Gate Congestion & Fake Passes', p: 'Manual paper registers at auditorium doors cause long queues & ticket duplication.' },
      { h: '⚔️ Hackathon Nightmares', p: 'Team formation friction, lost submissions & manual judging sheets → delayed results & disputes.' },
      { h: '📜 Certificate & Reporting Burden', p: 'Organizers spend weeks emailing Canva certificates; faculty spend weeks compiling NAAC/IQAC reports.' }
    ],
    [
      { h: '🎓 Student / Attendee Portal', p: 'One-click event discovery, team builder, digital ticket pass & verifiable certificate vault.' },
      { h: '🎪 Organizer & Faculty Coordinator Portal', p: 'Event builder, ticketing tiers, coupon engine, broadcast center & live attendance scanner.' },
      { h: '⚖️ Hackathon Judging & Mentor Console', p: 'Multi-criteria weighted scoring rubrics, submission inspection & live dynamic leaderboards.' },
      { h: '🏛️ College Admin & Dean Oversight Panel', p: 'Event approval workflows, cross-department analytics, asset libraries & system audit logs.' }
    ],
    [
      { h: '🔐 Secure & Effortless Auth', p: 'Google OAuth, GitHub OAuth, OTP login & 2FA with session management.' },
      { h: '🎫 Ticket & Pass Wallet', p: 'Instant PDF tickets with QR codes, add-to-calendar (.ics) & live pass status.' },
      { h: '🤝 Solo Hacker & Team Matchmaking', p: 'Find teammates by skill tag — Frontend, AI/ML, Backend, UI/UX.' },
      { h: '🏆 Gamification & Ranks', p: 'Experience badges, departmental ranks & verified activity logs.' },
      { h: '🌐 Social & Community', p: 'Follow clubs, bookmark events, verified testimonials & upvote showcase projects.' },
      { h: '👤 Public Profile · /u/username', p: 'On-platform portfolio of badges, ranks & verified achievements.' }
    ],
    [
      { h: '👥 Dynamic Team Workspace', p: 'Create teams, share auto-generated secure invite codes & assign project roles.' },
      { h: '📦 Submission Engine', p: 'Submit GitHub repos, live demo URLs, pitch videos, tech-stack tags & docs.' },
      { h: '⚖️ Digital Judging Portal', p: 'Multi-track judge assignment & weighted rubrics — Innovation 30% · Feasibility 30% · Code Quality 20% · Presentation 20%.' },
      { h: '🥇 Real-Time Dynamic Leaderboard', p: 'Auto-computed rankings, tie-breaker handling & instant podium reveals — zero grading disputes.' }
    ],
    [
      { h: '📷 In-Browser QR Scanner', p: 'Works on any smartphone camera or webcam with sub-second verification — no special hardware.' },
      { h: '⏱️ Multi-Session Attendance', p: 'Track checkpoints across Day 1 Inauguration, Lunch, Day 2 Hackathon & Valedictory.' },
      { h: '📶 PWA & Offline Fallback Protocol', p: 'Campus Wi-Fi fails? Scanner caches the guest list in encrypted local storage (IndexedDB), validates offline, auto-syncs on reconnect.' },
      { h: '⚡ Zero Gate Delays', p: 'Sub-second verification, zero duplicate check-ins — gates never stop moving.' }
    ],
    [
      { h: '🎨 Visual Certificate Studio', p: 'Custom templates for Participants, Winners, Runners-Up, Mentors & Volunteers with {{name}}, {{event}}, {{date}}, {{rank}} placeholders.' },
      { h: '⚡ Bulk Queue Dispatch', p: 'BullMQ + Puppeteer workers generate hundreds of high-res PDF/PNG certificates in the background.' },
      { h: '🔍 Tamper-Proof QR Verification', p: 'Every certificate carries a unique cryptographic QR linking to the official /verify page — zero forgery.' },
      { h: '💼 Placement & LinkedIn Ready', p: 'One-click sharing to LinkedIn to boost student employability & showcase college achievements.' }
    ],
    [
      { h: '📄 One-Click Official Export', p: 'Standardized institutional report templates in DOCX & PDF.' },
      { h: '📊 Pre-Populated Sections', p: 'Executive summary, participant lists with timestamps, speaker profiles, photo collage, feedback & financial balance sheet.' },
      { h: '🤖 AI-Assisted Drafting', p: 'Groq LPU + Gemini compile executive summaries & outcome analysis from raw event data.' },
      { h: '🏛️ Accreditation Compliant', p: 'Direct adherence to NAAC Criterion V, IQAC & AICTE/NIRF documentation guidelines.' }
    ],
    [
      { h: '💳 Integrated Razorpay Gateway', p: 'UPI (GPay, PhonePe, Paytm), Credit/Debit Cards & Net Banking.' },
      { h: '🎟️ Flexible Ticketing & Coupons', p: 'Free tickets, multi-tier pricing (Early Bird · Standard · VIP) & discount coupon management.' },
      { h: '🧾 Automated GST & Invoicing', p: 'Instant digital invoices & payment receipts issued to attendees.' },
      { h: '🛡️ Zero Cash Leakage & Audit Logs', p: 'Tamper-proof transaction logs for faculty & accounts, with automated bulk refunds.' }
    ],
    [
      { h: '📢 Email Broadcast Engine', p: 'Segmented 1-click updates to registered attendees, shortlisted teams or everyone.' },
      { h: '📁 Platform Asset Library', p: 'Central repository for college banners, department logos, sponsor decks & guidelines.' },
      { h: '📈 Live Analytics & Intelligence', p: 'Registration velocity, drop-off rates, revenue & venue capacity in real time.' },
      { h: '🤖 24/7 AI Attendee Assistant', p: 'Low-latency Groq + Gemini chatbot answering schedule, venue, rules & contact FAQs.' }
    ],
    [
      { h: '⚡ Frontend', p: 'Next.js 16 (App Router, SSR/SSG) · React 19 · Tailwind CSS · Zustand · Framer Motion · PWA offline.' },
      { h: '🛠️ Backend', p: 'Node.js 22 · Express.js · TypeScript · Mongoose ODM · REST API.' },
      { h: '🔄 Async Processing', p: 'Redis & BullMQ workers — Puppeteer PDF rendering, certificate generation, Nodemailer/Resend bulk mailers.' },
      { h: '🔒 Enterprise Security', p: 'JWT + refresh rotation, HTTP-only cookies, RBAC & CSRF, rate limiting, Helmet, Zod validation, audit logs & Sentry.' }
    ],
    [
      { h: '🛡️ Event Approval Workflow', p: 'Club events require faculty/Dean review before publishing — no unauthorized listings.' },
      { h: '🏢 Club & Organizer Verification', p: 'Verify clubs, assign organizer permissions & audit coordinator actions.' },
      { h: '📊 Cross-Department Insights', p: 'Compare participation across CSE, IT, ECE, Mech & MBA; annual activity summaries.' },
      { h: '🎛️ Platform Feature Flags', p: 'Enable/disable payments, hackathons & public signups dynamically via backend toggles.' }
    ],
    [
      { h: '1️⃣ Student Flow', p: 'Browse events → Register in 1 click → QR Pass appears in wallet.' },
      { h: '2️⃣ Hackathon Flow', p: 'Form team with invite code → Submit demo URL & GitHub repo.' },
      { h: '3️⃣ Gate Scanner', p: 'Organizer mobile camera scans student QR → sub-second attendance confirmation.' },
      { h: '4️⃣ Admin / Faculty Flow', p: 'One-click certificate release with QR verification → one-click NAAC/IQAC report download.' }
    ],
    [
      { h: '📅 Phase 1 — Pilot Deployment', p: 'Weeks 1–2 · Deploy for 2 departmental workshops; validate gate scanning & certificate dispatch.' },
      { h: '📅 Phase 2 — Club & Faculty Onboarding', p: 'Weeks 3–4 · 30-min training for club leads, societies (IEEE, GDSC, ACM, Cultural) & coordinators.' },
      { h: '📅 Phase 3 — Campus-Wide Launch', p: 'Month 2 · Official event portal linked on the college website & student app.' },
      { h: '🎯 Our Request Today', p: 'Seeking official approval from College Administration to launch the Phase 1 Pilot.' }
    ],
    [
      { h: '"What if the internet fails during a 1,000-person fest?"', p: 'Offline PWA Protocol — encrypted local cache, sub-0.5s validation, auto-sync on reconnect.' },
      { h: '"Can a student create a fake or inappropriate event?"', p: 'Strict RBAC — only verified club organizers can create, and events need faculty/admin approval.' },
      { h: '"How does this help NAAC / NBA accreditation?"', p: 'AI report generator exports structured DOCX/PDF formatted for NAAC Criterion V.' },
      { h: '"Is student data secure and private?"', p: 'JWT with refresh rotation, 2FA, encrypted hashing, role-based queries & rate-limiting.' },
      { h: '"Will the server crash if 2,000 students register at once?"', p: 'Async BullMQ/Redis workers handle PDFs, certificates & mail — API stays light.' }
    ],
    null
  ];

  var CL = {
    bg:'050505', panel:'0F0F14', panel2:'19191E', border:'26262D',
    text:'F5F5F5', muted:'9CA3AF', dim:'71717A',
    red:'DC2626', red2:'EF4444', red3:'F87171', gold:'FFD700', green:'22C55E'
  };
  var FH='Chakra Petch', FB='Rajdhani', FM='JetBrains Mono';
  var W=13.333, H=7.5;

  function clean(s){ return String(s||'').replace(/<[^>]+>/g,'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;|&apos;/g,"'").replace(/&nbsp;/g,' ').replace(/\s+/g,' ').trim(); }

  function deckData(){
    var out=[];
    document.querySelectorAll('.slide').forEach(function(sec){
      var k=(sec.querySelector('.kicker')||{}).textContent||'';
      var t=(sec.querySelector('.s-title')||{}).textContent||sec.dataset.title||'';
      var s=(sec.querySelector('.s-sub')||{}).textContent||'';
      out.push({ title: sec.dataset.title||'', notes: sec.dataset.notes||'', kicker: clean(k), sTitle: clean(t), sSub: clean(s) });
    });
    return out;
  }

  function brandTop(slide){
    slide.addShape('roundRect',{x:0.45,y:0.38,w:0.42,h:0.42,rectRadius:0.08,fill:{color:CL.red},line:{type:'none'}});
    slide.addText('N',{x:0.45,y:0.38,w:0.42,h:0.42,align:'center',valign:'middle',fontFace:FH,fontSize:16,bold:true,color:'FFFFFF'});
    slide.addText([{text:'NX-913  ',options:{fontFace:FH,bold:true,color:CL.text}},{text:'CAMPUS OS',options:{fontFace:FB,color:CL.dim,fontSize:11}}],{x:0.97,y:0.4,w:3.2,h:0.38,valign:'middle'});
    slide.addText('PITCH DECK',{x:W-2.6,y:0.44,w:2.15,h:0.3,align:'right',fontFace:FM,fontSize:9,color:CL.dim,charSpacing:2});
  }
  function footer(slide,idx,total){
    slide.addShape('rect',{x:0,y:H-0.09,w:W*((idx+1)/total),h:0.05,fill:{color:CL.red}});
    slide.addText('NX-913 · CAMPUS EVENT & HACKATHON OS',{x:0.45,y:H-0.55,w:6,h:0.3,fontFace:FB,fontSize:9,color:CL.dim,charSpacing:1.5});
    slide.addText(String(idx+1).padStart(2,'0')+' / '+String(total).padStart(2,'0'),{x:W-1.6,y:H-0.55,w:1.15,h:0.3,align:'right',fontFace:FM,fontSize:10,color:CL.muted});
  }
  function header(slide,d,idx){
    var kicker=d.kicker||('SLIDE '+String(idx+1).padStart(2,'0'));
    slide.addShape('roundRect',{x:0.45,y:1.05,w:1.25,h:0.38,rectRadius:0.19,fill:{color:'1A0D0D'},line:{color:'7F1D1D',width:1}});
    slide.addText(String(kicker).toUpperCase(),{x:0.45,y:1.05,w:1.25,h:0.38,align:'center',valign:'middle',fontFace:FM,fontSize:9,color:CL.red3,charSpacing:1});
    slide.addText(d.sTitle||d.title,{x:0.42,y:1.5,w:W-0.9,h:0.9,valign:'middle',fontFace:FH,fontSize:34,bold:true,color:CL.text});
    if(d.sSub){ slide.addText(d.sSub,{x:0.45,y:2.32,w:W-1.8,h:0.45,valign:'top',fontFace:FB,fontSize:15,color:CL.muted}); }
  }
  function card(slide,x,y,w,h,item,accent,compact){
    var hs=compact?12.5:14.5, ps=compact?11:12.5, hy=compact?0.07:0.1, py=compact?0.44:0.58;
    slide.addShape('roundRect',{x:x,y:y,w:w,h:h,rectRadius:0.1,fill:{color:CL.panel},line:{color:CL.border,width:1}});
    slide.addShape('rect',{x:x,y:y+0.12,w:0.05,h:h-0.24,fill:{color:accent||CL.red}});
    slide.addText(item.h,{x:x+0.22,y:y+hy,w:w-0.4,h:0.42,valign:'top',fontFace:FH,fontSize:hs,bold:true,color:CL.text,breakLine:false});
    if(item.p){ slide.addText(item.p,{x:x+0.22,y:y+py,w:w-0.4,h:h-py-0.12,valign:'top',fontFace:FB,fontSize:ps,color:CL.muted,lineSpacing:18}); }
  }
  function bodyGrid(slide,items,accent){
    var top=2.95,bottom=H-0.75,n=items.length,cols=n===5?2:3,rows=Math.ceil(n/cols),gap=0.22;
    var colW=(W-0.9-gap*(cols-1))/cols,rowH=(bottom-top-gap*(rows-1))/rows;
    var compact=rowH<1.3;
    items.forEach(function(it,i){
      var r=Math.floor(i/cols),c=i%cols;
      card(slide,0.45+c*(colW+gap),top+r*(rowH+gap),colW,rowH,it,accent,compact);
    });
  }
  function framedShot(slide,dataUri,x,y,w,h,url){
    slide.addShape('roundRect',{x:x,y:y,w:w,h:h,rectRadius:0.09,fill:{color:'0C0C10'},line:{color:CL.border2,width:1}});
    slide.addShape('rect',{x:x+0.06,y:y+0.06,w:w-0.12,h:0.34,fill:{color:'14141A'},line:{type:'none'}});
    ['7F1D1D','B45309','166534'].forEach(function(c,i){
      slide.addShape('ellipse',{x:x+0.14+i*0.2,y:y+0.155,w:0.12,h:0.12,fill:{color:c},line:{type:'none'}});
    });
    slide.addShape('roundRect',{x:x+0.72,y:y+0.13,w:w-1.0,h:0.2,rectRadius:0.1,fill:{color:'0A0A0E'},line:{color:CL.border,width:0.5}});
    slide.addText(url||'nx-913.com',{x:x+0.78,y:y+0.13,w:w-1.1,h:0.2,align:'center',valign:'middle',fontFace:FM,fontSize:8,color:CL.dim});
    slide.addImage({data:'data:image/jpeg;base64,'+dataUri,x:x+0.06,y:y+0.46,w:w-0.12,h:h-0.56});
  }
  function titleSlide(slide,d,idx,total){
    brandTop(slide);
    slide.addShape('ellipse',{x:-1.5,y:-1.2,w:7,h:7,fill:{color:CL.red,transparency:88},line:{type:'none'}});
    slide.addShape('roundRect',{x:0.85,y:2.0,w:2.2,h:0.46,rectRadius:0.23,fill:{color:'120B0B'},line:{color:'7F1D1D',width:1}});
    slide.addText('N X - 9 1 3',{x:0.85,y:2.0,w:2.2,h:0.46,align:'center',valign:'middle',fontFace:FM,fontSize:11,color:CL.red3,charSpacing:1.5});
    slide.addText('NX-913',{x:0.82,y:2.55,w:6.6,h:1.4,align:'left',valign:'middle',fontFace:FH,fontSize:66,bold:true,color:CL.red2,charSpacing:1});
    slide.addText('The Next-Generation Campus Event & Hackathon Operating System',{x:0.85,y:3.95,w:7.1,h:0.9,align:'left',fontFace:FH,fontSize:20,bold:true,color:CL.text,lineSpacing:28});
    slide.addText('From Registration to NAAC Reports — In One Click',{x:0.85,y:5.0,w:7.1,h:0.4,align:'left',fontFace:FB,fontSize:15,color:CL.gold});
    slide.addText('Unified Student Engagement · Hackathon Lifecycle · Academic Compliance',{x:0.85,y:5.5,w:7.1,h:0.4,align:'left',fontFace:FB,fontSize:13,color:CL.muted});
    var chips=['🎓 Student Portal','🎪 Organizer Suite','⚖️ Judging Console','🏛️ Admin Oversight'];
    chips.forEach(function(c,i){
      var cw=1.62,gap=0.18;
      slide.addShape('roundRect',{x:0.85+i*(cw+gap),y:6.05,w:cw,h:0.42,rectRadius:0.21,fill:{color:CL.panel2},line:{color:CL.border,width:1}});
      slide.addText(c,{x:0.85+i*(cw+gap),y:6.05,w:cw,h:0.42,align:'center',valign:'middle',fontFace:FB,fontSize:10.5,bold:true,color:CL.text});
    });
    framedShot(slide,SHOTS.home,8.35,1.35,4.4,5.35,'https://nx-913.com');
    footer(slide,idx,total);
  }
  function thanksSlide(slide,d,idx,total){
    brandTop(slide);
    slide.addShape('ellipse',{x:W-5.5,y:-1.5,w:7,h:7,fill:{color:CL.red,transparency:90},line:{type:'none'}});
    slide.addText('Thank You',{x:0,y:2.2,w:W,h:1.5,align:'center',valign:'middle',fontFace:FH,fontSize:66,bold:true,color:CL.text});
    slide.addText('NX-913 — the unified operating system for every campus event, hackathon & accreditation report.',{x:W/2-4.2,y:3.95,w:8.4,h:0.55,align:'center',fontFace:FB,fontSize:16,color:CL.muted});
    slide.addText('“From Registration to NAAC Reports — In One Click.”',{x:W/2-3.5,y:4.6,w:7,h:0.4,align:'center',fontFace:FB,fontSize:13.5,color:CL.gold});
    var chips=['🎟️ Events & Passes','🏆 Hackathons','📜 Certificates','📄 NAAC Reports'];
    chips.forEach(function(c,i){
      var cw=2.35,gap=0.3,totalW=chips.length*cw+(chips.length-1)*gap,x=W/2-totalW/2+i*(cw+gap);
      slide.addShape('roundRect',{x:x,y:5.3,w:cw,h:0.5,rectRadius:0.25,fill:{color:CL.panel2},line:{color:CL.border,width:1}});
      slide.addText(c,{x:x,y:5.3,w:cw,h:0.5,align:'center',valign:'middle',fontFace:FB,fontSize:12,bold:true,color:CL.text});
    });
    slide.addText('QUESTIONS & DISCUSSION — OPEN FLOOR',{x:W/2-3,y:6.15,w:6,h:0.4,align:'center',fontFace:FM,fontSize:11,color:CL.red3,charSpacing:2});
    footer(slide,idx,total);
  }
  function sideShotSlide(slide,d,idx,total,items,shot,url,accent){
    brandTop(slide);
    header(slide,d,idx);
    var leftW=7.7,top=3.0,bottom=H-0.75,n=items.length,cols=2,rows=Math.ceil(n/cols),gap=0.2;
    var colW=(leftW-gap)/2,rowH=(bottom-top-gap*(rows-1))/rows,compact=rowH<1.35;
    items.forEach(function(it,i){
      var r=Math.floor(i/cols),c=i%cols;
      card(slide,0.45+c*(colW+gap),top+r*(rowH+gap),colW,rowH,it,accent,compact);
    });
    framedShot(slide,shot,8.4,2.75,4.35,4.15,url);
  }

  window.downloadPPT = function(){
    var slides=[];
    var data=deckData(),total=data.length;
    var PptxGenJS=window.PptxGenJS;
    if(!PptxGenJS){ alert('PPT engine not loaded — check your internet connection, then try again.'); return; }
    var pptx=new PptxGenJS();
    pptx.defineLayout({name:'WIDE',width:W,height:H});
    pptx.layout='WIDE';
    pptx.author='NX-913';
    pptx.title='NX-913 — Campus Event & Hackathon Operating System';
    data.forEach(function(d,i){
      var slide=pptx.addSlide();
      slide.background={color:CL.bg};
      var body=BODIES[i];
      if(i===0){ titleSlide(slide,d,i,total); }
      else if(i===total-1){ thanksSlide(slide,d,i,total); }
      else if(i===3){ sideShotSlide(slide,d,i,total,body,SHOTS.events,'https://nx-913.com/events'); }
      else if(i===4){ sideShotSlide(slide,d,i,total,body,SHOTS.detail,'https://nx-913.com/hackathons'); }
      else if(i===12){ sideShotSlide(slide,d,i,total,body,SHOTS.hackathons,'https://nx-913.com/hackathons'); }
      else { brandTop(slide); header(slide,d,i); if(body) bodyGrid(slide,body); }
      if(d.notes) slide.addNotes(d.notes);
    });
    return pptx.writeFile({fileName:'NX-913_Pitch_Deck.pptx'});
  };

  /* load pptxgenjs from CDN lazily, then expose the button action */
  window.loadPPTEngine = function(cb){
    if(window.PptxGenJS){ cb&&cb(); return; }
    var s=document.createElement('script');
    s.src='https://cdn.jsdelivr.net/npm/pptxgenjs@4.0.1/dist/pptxgen.bundle.min.js';
    s.onload=function(){ cb&&cb(); };
    s.onerror=function(){ cb&&cb(new Error('cdn')); };
    document.head.appendChild(s);
  };
})();
