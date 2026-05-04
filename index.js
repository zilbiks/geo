const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'Starptautiskais tūrisms: VIP — Albānija';

// Color palette
const C = {
  black: "0A0A0A",
  darkGray: "111111",
  gray1: "1A1A1A",
  gray2: "242424",
  gray3: "2E2E2E",
  gray4: "3A3A3A",
  red: "CC0000",
  redDark: "990000",
  redLight: "E53333",
  white: "FFFFFF",
  offWhite: "F0F0F0",
  silver: "AAAAAA",
  lightGray: "CCCCCC",
};

const makeShadow = () => ({ type: "outer", blur: 12, offset: 3, angle: 135, color: "000000", opacity: 0.5 });

// ==============================
// SLIDE 1: TITLE
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.black };

  // Big red accent bar left
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.red }, line: { color: C.red } });

  // Country silhouette / decorative element - big faint text
  s.addText("ALBANIA", {
    x: 3.5, y: 0.5, w: 6, h: 2, fontSize: 110, bold: true, color: "1A1A1A",
    fontFace: "Arial Black", align: "right", valign: "middle", margin: 0
  });

  // Main title
  s.addText("Starptautiskais tūrisms: VIP", {
    x: 0.5, y: 0.5, w: 7, h: 0.7, fontSize: 18, bold: false, color: C.silver,
    fontFace: "Calibri", align: "left", charSpacing: 4, margin: 0
  });
  s.addText("ALBĀNIJA", {
    x: 0.5, y: 1.1, w: 7, h: 1.6, fontSize: 80, bold: true, color: C.white,
    fontFace: "Arial Black", align: "left", margin: 0
  });

  // Red accent line
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.75, w: 2.5, h: 0.05, fill: { color: C.red }, line: { color: C.red } });

  s.addText("Projekts 12. klasei  -  20.04.–24.04.2026.", {
    x: 0.5, y: 2.9, w: 7, h: 0.4, fontSize: 14, color: C.silver,
    fontFace: "Calibri", align: "left", margin: 0
  });

  // 3 stages
  const stages = ["1. POSMS\nVIEDOKĻI", "2. POSMS\nIZVĒRTĒŠANA", "3. POSMS\nPĒTĪJUMS"];
  stages.forEach((t, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.2, w: 2.8, h: 0.9,
      fill: { color: C.gray2 }, line: { color: C.gray3, width: 1 }
    });
    // Red top accent
    s.addShape(pres.shapes.RECTANGLE, { x, y: 4.2, w: 2.8, h: 0.05, fill: { color: C.red }, line: { color: C.red } });
    const lines = t.split("\n");
    s.addText([
      { text: lines[0], options: { bold: true, color: C.red, fontSize: 10, breakLine: true } },
      { text: lines[1], options: { bold: true, color: C.white, fontSize: 13 } }
    ], { x, y: 4.25, w: 2.8, h: 0.85, align: "center", valign: "middle", fontFace: "Calibri" });
  });

  // Latvian cities
  s.addText("Balvi - Daugavpils - Ventspils - Liepāja", {
    x: 0.5, y: 5.15, w: 9, h: 0.35, fontSize: 11, color: C.gray4,
    fontFace: "Calibri", align: "left", margin: 0
  });
}

// ==============================
// SLIDE 2: POSMS 1 — SECTION
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.red };

  s.addText("01", {
    x: 0, y: -0.3, w: 10, h: 4, fontSize: 200, bold: true, color: "BB0000",
    fontFace: "Arial Black", align: "right", valign: "middle", margin: 0
  });

  s.addText("POSMS", {
    x: 0.7, y: 0.6, w: 5, h: 0.6, fontSize: 16, color: "FFAAAA",
    fontFace: "Calibri", charSpacing: 6, align: "left", margin: 0
  });
  s.addText("VIEDOKĻI", {
    x: 0.7, y: 1.1, w: 6, h: 1.4, fontSize: 72, bold: true, color: C.white,
    fontFace: "Arial Black", align: "left", margin: 0
  });
  s.addText("Intervijas ar cilvēkiem par starptautisko tūrisma pieredzi", {
    x: 0.7, y: 2.6, w: 7, h: 0.6, fontSize: 16, color: "FFCCCC",
    fontFace: "Calibri", align: "left", margin: 0
  });
}

// ==============================
// SLIDE 3: INTERVIEW QUESTIONS
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("1. POSMS  -  INTERVIJAS JAUTĀJUMI", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const qs = [
    "Kuru valsti esat apmeklējuši kā tūrists un kāpēc izvēlējāties tieši to?",
    "Kādas bija Jūsu gaidas pirms ceļojuma un vai tās piepildījās?",
    "Kādi ir trīs atmiņā palikušie iespaidīgākie piedzīvojumi vai vietas?",
    "Kādas grūtības vai neērtības radās ceļojuma laikā?",
    "Ko ieteiktu citiem, kuri plāno doties uz šo valsti?",
    "Cik ilgs bija ceļojums un kā nokļuvāt galamērķī?",
    "Kā novērtētu vietējos iedzīvotājus un viņu attieksmi pret tūristiem?",
    "Vai plānojat atkārtoti apmeklēt šo valsti? Kāpēc?",
  ];

  const colW = 4.4;
  qs.forEach((q, i) => {
    const col = i < 4 ? 0 : 1;
    const row = i % 4;
    const x = 0.4 + col * 5.0;
    const y = 1.05 + row * 1.1;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: colW, h: 0.95,
      fill: { color: C.gray2 }, line: { color: C.gray3, width: 1 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 0.95, fill: { color: C.red }, line: { color: C.red } });
    s.addText([
      { text: `${i + 1}.  `, options: { bold: true, color: C.red, fontSize: 12 } },
      { text: q, options: { color: C.lightGray, fontSize: 11 } }
    ], { x: x + 0.1, y, w: colW - 0.15, h: 0.95, valign: "middle", fontFace: "Calibri" });
  });
}

// ==============================
// SLIDE 4: INTERVIEW 1
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("1. POSMS  -  INTERVIJA Nr. 1", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  // Subtitle
  s.addText("Albānija – Tirana un Berat  |  2022. gads, 10 dienas", {
    x: 0.4, y: 0.95, w: 9, h: 0.45, fontSize: 13, color: C.silver,
    fontFace: "Calibri", italic: true, margin: 0
  });

  const rows = [
    ["APMEKLĒTĀ VALSTS", "Albānija – Tirana un Berat (2022. gadā, 10 dienas)"],
    ["GAIDAS", "Gaidīja neizpētītu Eiropas valsti ar unikālu kultūru. Gaidas pilnībā piepildījās."],
    ["IESPAIDĪGĀKĀS VIETAS", "Berat vecpilsēta - Skadar ezers - Skanderbega laukums Tiranā"],
    ["GRŪTĪBAS", "Ceļu infrastruktūra lauku apvidos – daļa maršrutu nav asfaltēta"],
    ["IETEIKUMI", "Byrek un tavë kosi ir obligāti! Ņemt līdzi skaidru naudu."],
    ["TRANSPORTS", "Rīga -> Vīne -> Tirana  (Austrian Airlines, ~5 h)"],
    ["ATKĀRTOTS APMEKLĒJUMS", "Noteikti! Vēlas apmeklēt Albānijas Alpes (Theth) un Xhamil pludmales."],
  ];

  rows.forEach((r, i) => {
    const y = 1.45 + i * 0.58;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.52,
      fill: { color: i % 2 === 0 ? C.gray2 : C.gray1 }, line: { color: C.gray2 }
    });
    s.addText(r[0], { x: 0.5, y, w: 2.5, h: 0.52, fontSize: 9.5, bold: true, color: C.red, fontFace: "Calibri", valign: "middle", charSpacing: 1 });
    s.addText(r[1], { x: 3.0, y, w: 6.5, h: 0.52, fontSize: 11, color: C.offWhite, fontFace: "Calibri", valign: "middle" });
  });
}

// ==============================
// SLIDE 5: INTERVIEW 2
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("1. POSMS  -  INTERVIJA Nr. 2", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  s.addText("Albānija + Melnkalne  |  2023. gads, 12 dienas", {
    x: 0.4, y: 0.95, w: 9, h: 0.45, fontSize: 13, color: C.silver,
    fontFace: "Calibri", italic: true, margin: 0
  });

  const rows = [
    ["APMEKLĒTĀ VALSTS", "Albānija un Melnkalne kombinētā tura veidā (2023. gadā, 12 dienas)"],
    ["GAIDAS", "Interesēja Balkānu vēsture un pieejamas cenas. Albānija bija lētākā Eiropas valsts."],
    ["IESPAIDĪGĀKĀS VIETAS", "Butrint UNESCO drupas - Xhamil kristālskaidrais ūdens - Gjirokastra cietoksnis"],
    ["GRŪTĪBAS", "Valodas barjera lauku rajonos – angļu valoda zināma tikai jauniešiem."],
    ["IETEIKUMI", "Nomāt auto – vienīgais veids, kā sasniegt tālākos punktus. Albānieši ir ļoti draudzīgi."],
    ["TRANSPORTS", "Rīga -> Helsinki -> Tirana  (Finnair, ~6 h)"],
    ["ATKĀRTOTS APMEKLĒJUMS", "Jā! Vēlas apmeklēt arī Vlorë un dienvidu piekrasti pavasarī."],
  ];

  rows.forEach((r, i) => {
    const y = 1.45 + i * 0.58;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.52,
      fill: { color: i % 2 === 0 ? C.gray2 : C.gray1 }, line: { color: C.gray2 }
    });
    s.addText(r[0], { x: 0.5, y, w: 2.5, h: 0.52, fontSize: 9.5, bold: true, color: C.red, fontFace: "Calibri", valign: "middle", charSpacing: 1 });
    s.addText(r[1], { x: 3.0, y, w: 6.5, h: 0.52, fontSize: 11, color: C.offWhite, fontFace: "Calibri", valign: "middle" });
  });
}

// ==============================
// SLIDE 6: INTERVIEW CONCLUSIONS
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("1. POSMS  -  SECINĀJUMI NO INTERVIJĀM", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const cards = [
    ["NEIZPĒTĪTS DĀRGAKMENS", "Abiem intervējamajiem Albānija bija pozitīvas pārsteiguma pilna pieredze — vēl nav pārsātināta ar tūristiem."],
    ["TRANSPORTS", "Rīga -> Tirana ir sasniedzama ar 1 pārsēšanos caur Vīni vai Helsinkiem (~5–6 h kopā)."],
    ["CENAS", "Ēdināšana, nakšņošana un ieejas biļetes ir ievērojami lētākas nekā Rīgā vai Rietumeiropā."],
    ["GRŪTĪBAS", "Angļu valoda nav pieejama lauku rajonos; daļa ceļu nav asfaltēti."],
    ["IEDZĪVOTĀJI", "Abi intervējamie uzsvēra albāniešu draudzīgumu un atvērtību tūristiem."],
  ];

  cards.forEach((c, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const isFull = i === 4;
    const x = isFull ? 0.4 : 0.4 + col * 4.7;
    const w = isFull ? 9.2 : 4.4;
    const y = 1.1 + row * 1.55;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 1.35,
      fill: { color: C.gray2 }, line: { color: C.gray3, width: 1 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.35, fill: { color: C.red }, line: { color: C.red } });
    s.addText(c[0], { x: x + 0.15, y: y + 0.1, w: w - 0.2, h: 0.35, fontSize: 10, bold: true, color: C.red, fontFace: "Calibri", charSpacing: 1 });
    s.addText(c[1], { x: x + 0.15, y: y + 0.42, w: w - 0.25, h: 0.85, fontSize: 11.5, color: C.lightGray, fontFace: "Calibri", valign: "top" });
  });
}

// ==============================
// SLIDE 7: POSMS 2 — SECTION
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.gray1 };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.red }, line: { color: C.red } });

  s.addText("02", {
    x: 0, y: -0.3, w: 10, h: 4, fontSize: 200, bold: true, color: "1E1E1E",
    fontFace: "Arial Black", align: "right", valign: "middle", margin: 0
  });

  s.addText("POSMS", {
    x: 0.7, y: 0.6, w: 5, h: 0.6, fontSize: 16, color: C.silver,
    fontFace: "Calibri", charSpacing: 6, align: "left", margin: 0
  });
  s.addText("IZVĒRTĒŠANA", {
    x: 0.7, y: 1.1, w: 8, h: 1.4, fontSize: 60, bold: true, color: C.white,
    fontFace: "Arial Black", align: "left", margin: 0
  });
  s.addText("Latvijas tūrisma galamērķi — ārvalstu tūrista skatījumā", {
    x: 0.7, y: 2.6, w: 7, h: 0.6, fontSize: 16, color: C.silver,
    fontFace: "Calibri", align: "left", margin: 0
  });
  s.addText("Daugavpils - Ventspils - Liepāja", {
    x: 0.7, y: 3.2, w: 7, h: 0.5, fontSize: 14, color: C.red,
    fontFace: "Calibri", bold: true, align: "left", margin: 0
  });
}

// ==============================
// SLIDE 8: EVALUATION CRITERIA
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("2. POSMS  -  VĒRTĒŠANAS KRITĒRIJI", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const criteria = [
    ["01", "PIEEJAMĪBA", "Transports, ceļi, attālums no Rīgas"],
    ["02", "APSKATES OBJEKTI", "Kultūrvēsturiski, dabas un arhitektoniski objekti"],
    ["03", "NAKŠŅOŠANA", "Viesnīcu skaits, kategorijas, cenas"],
    ["04", "UZTURS", "Restorāni, vietējā virtuve, pieejamība"],
    ["05", "AKTĪVĀ ATPŪTA", "Sports, daba, izklaide"],
    ["06", "INFO PIEEJAMĪBA", "Tūrisma centri, angļu valoda"],
    ["07", "CENA / KVALITĀTE", "Vai pakalpojumi atbilst cenai?"],
  ];

  criteria.forEach((cr, i) => {
    const col = i < 4 ? 0 : 1;
    const row = i < 4 ? i : i - 4;
    const x = 0.4 + col * 5.0;
    const y = 1.05 + row * 1.1;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.4, h: 0.95,
      fill: { color: C.gray2 }, line: { color: C.gray3 }
    });

    // Number circle
    s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: y + 0.2, w: 0.52, h: 0.52, fill: { color: C.red }, line: { color: C.red } });
    s.addText(cr[0], { x: x + 0.12, y: y + 0.2, w: 0.52, h: 0.52, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle", fontFace: "Calibri" });

    s.addText(cr[1], { x: x + 0.75, y: y + 0.08, w: 3.5, h: 0.38, fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", valign: "bottom" });
    s.addText(cr[2], { x: x + 0.75, y: y + 0.45, w: 3.5, h: 0.4, fontSize: 10, color: C.silver, fontFace: "Calibri", valign: "top" });
  });

  // Last one is alone in col 1 row 3 (index 6 = col 1, row 3 of the 4-spot grid)
  // Already handled above
}

// ==============================
// SLIDE 9: COMPARISON TABLE
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("2. POSMS  -  LATVIJAS GALAMĒRĶU SALĪDZINĀJUMS", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const headers = [
    [{ text: "KRITĒRIJS", options: { bold: true, color: C.red, fontSize: 10 } }],
    [{ text: "DAUGAVPILS", options: { bold: true, color: C.white, fontSize: 10 } }],
    [{ text: "VENTSPILS", options: { bold: true, color: C.white, fontSize: 10 } }],
    [{ text: "LIEPĀJA", options: { bold: true, color: C.white, fontSize: 10 } }],
  ];

  const rows = [
    ["Pieejamība", "Autobuss/vilciens ~3h", "Autobuss ~3h", "Autobuss/vilciens ~2.5h"],
    ["Apskates objekti", "Cietoksnis, Rothko centrs", "Kūrorts, bērnu parks", "Karosta, ezers, mūzika"],
    ["Nakšņošana", "Pieejamas cenas", "Viesnīcas kūrortā", "Laba cena/kvalitāte"],
    ["Uzturs", "Slāvu/latviešu virtuve", "Zivju restorāni", "Kafejnīcas, live mūzika"],
    ["Aktīvā atpūta", "Velo, cietokšņa tūres", "Pludmale, ūdens parks", "Surf, pludmale, ezers"],
    ["KOPVĒRTĒJUMS", "7.5 / 10", "8.5 / 10", "8.0 / 10"],
  ];

  const colW = [2.5, 2.2, 2.2, 2.2];
  const colX = [0.35, 2.9, 5.15, 7.4];

  // Header row
  colX.forEach((x, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.0, w: colW[ci], h: 0.5,
      fill: { color: ci === 0 ? C.gray3 : C.redDark }, line: { color: C.gray4 }
    });
    s.addText(headers[ci], { x, y: 1.0, w: colW[ci], h: 0.5, align: "center", valign: "middle", fontFace: "Calibri" });
  });

  rows.forEach((row, ri) => {
    const isLast = ri === rows.length - 1;
    const rowY = 1.52 + ri * 0.64;
    row.forEach((cell, ci) => {
      const bgColor = isLast ? C.red : (ri % 2 === 0 ? C.gray2 : C.gray1);
      const textColor = isLast ? C.white : (ci === 0 ? C.silver : C.lightGray);
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y: rowY, w: colW[ci], h: 0.58,
        fill: { color: bgColor }, line: { color: C.gray4 }
      });
      s.addText(cell, {
        x: colX[ci], y: rowY, w: colW[ci], h: 0.58,
        fontSize: isLast ? 13 : 10, bold: isLast, color: textColor,
        fontFace: "Calibri", align: "center", valign: "middle"
      });
    });
  });

  s.addText("UZVARĒTĀJS: VENTSPILS — vislabāk piemērota ārvalstu tūristam", {
    x: 0.35, y: 5.15, w: 9.3, h: 0.35, fontSize: 10, bold: true, color: C.red,
    fontFace: "Calibri", italic: true, margin: 0
  });
}

// ==============================
// SLIDE 10: VENTSPILS COLLAGE
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("2. POSMS  -  VENTSPILS — ĀRVALSTU TŪRISTA PERSPEKTĪVA", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  // Left column - positives
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.0, w: 4.4, h: 0.5, fill: { color: "1A3A1A" }, line: { color: "2A5A2A" } });
  s.addText("[OK]  POZITĪVIE ASPEKTI", { x: 0.35, y: 1.0, w: 4.4, h: 0.5, fontSize: 12, bold: true, color: "66DD66", fontFace: "Calibri", align: "center", valign: "middle" });

  const pros = [
    "Lielisks jūras kūrorts ar tīru pludmali",
    "Bērnu un ģimenes atrakciju parks",
    "Moderna infrastruktūra un promenāde",
    "Dabas takas netālu no centra",
    "Autentiska zvejas osta",
    "Sasniedzams ar sabiedrisko transportu",
  ];
  pros.forEach((p, i) => {
    const y = 1.55 + i * 0.54;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 4.4, h: 0.48, fill: { color: i % 2 === 0 ? C.gray2 : C.gray1 }, line: { color: C.gray3 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.06, h: 0.48, fill: { color: "44AA44" }, line: { color: "44AA44" } });
    s.addText(p, { x: 0.5, y, w: 4.1, h: 0.48, fontSize: 11, color: C.lightGray, fontFace: "Calibri", valign: "middle" });
  });

  // Right column - negatives
  s.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.0, w: 4.4, h: 0.5, fill: { color: "3A1A1A" }, line: { color: "5A2A2A" } });
  s.addText("[X]  NEGATĪVIE ASPEKTI", { x: 5.25, y: 1.0, w: 4.4, h: 0.5, fontSize: 12, bold: true, color: "FF6666", fontFace: "Calibri", align: "center", valign: "middle" });

  const cons = [
    "Sezonalitāte — aktīvs tikai vasarā",
    "Ziemā lielākā daļa atrakciju slēgta",
    "Ierobežots nakšņošanas piedāvājums",
    "Maz ēdināšanas vietu nakts stundās",
    "Angļu valoda tikai tūrisma objektos",
    "Bez auto pārvietošanās sarežģīta",
  ];
  cons.forEach((c, i) => {
    const y = 1.55 + i * 0.54;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.25, y, w: 4.4, h: 0.48, fill: { color: i % 2 === 0 ? C.gray2 : C.gray1 }, line: { color: C.gray3 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.25, y, w: 0.06, h: 0.48, fill: { color: C.red }, line: { color: C.red } });
    s.addText(c, { x: 5.4, y, w: 4.1, h: 0.48, fontSize: 11, color: C.lightGray, fontFace: "Calibri", valign: "middle" });
  });
}

// ==============================
// SLIDE 11: POSMS 3 — SECTION
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.black };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.red }, line: { color: C.red } });

  s.addText("03", {
    x: 0, y: -0.3, w: 10, h: 4, fontSize: 200, bold: true, color: "161616",
    fontFace: "Arial Black", align: "right", valign: "middle", margin: 0
  });

  s.addText("POSMS", {
    x: 0.7, y: 0.6, w: 5, h: 0.6, fontSize: 16, color: C.silver,
    fontFace: "Calibri", charSpacing: 6, align: "left", margin: 0
  });
  s.addText("PĒTĪJUMS", {
    x: 0.7, y: 1.1, w: 8, h: 1.4, fontSize: 72, bold: true, color: C.white,
    fontFace: "Arial Black", align: "left", margin: 0
  });
  s.addText("Tūrisma galamērķis: Albānija", {
    x: 0.7, y: 2.55, w: 7, h: 0.55, fontSize: 18, color: C.red, bold: true,
    fontFace: "Calibri", align: "left", margin: 0
  });
  s.addText("Iespējas un draudi  -  20.04.–24.04.2026.", {
    x: 0.7, y: 3.1, w: 7, h: 0.5, fontSize: 14, color: C.silver,
    fontFace: "Calibri", align: "left", margin: 0
  });
}

// ==============================
// SLIDE 12: GEOGRAPHY
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  ALBĀNIJA — ĢEOGRĀFISKAIS APRAKSTS", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const geoData = [
    ["NOVIETOJUMS", "Dienvideiropa, Balkānu pussala\n41°20' Z platums, 20°00' A garums"],
    ["ROBEŽVALSTIS", "Melnkalne - Serbija - Ziemeļmaķedonija - Grieķija"],
    ["JŪRAS", "Rietumos: Adrijas jūra - Dienvidrietumos: Jonijas jūra"],
    ["RELJEFS", "~70% kalnu — Albānijas Alpes ziemeļos. Piekrastes līdzenumi rietumos."],
    ["IEKŠĒJIE ŪDEŅI", "Skadar ezers (lielākais Balkānos) - Ohrid ezers - kalnu upes"],
    ["KLIMATS", "Vidusjūras (piekrastē): karsts, sauss vasarā\nKontinentāls: kalnos — vēsāks"],
    ["DABAS RESURSI", "Siltas pludmales - nacionālie parki - kalnu ainavas - seno pilsētu drupas"],
    ["TŪRISMA NOZĪME", "Albānija — viena no straujāk augošajiem galamērķiem Eiropā"],
  ];

  geoData.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.35 + col * 4.8;
    const y = 1.05 + row * 1.1;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.5, h: 1.0,
      fill: { color: C.gray2 }, line: { color: C.gray3 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.0, fill: { color: C.red }, line: { color: C.red } });
    s.addText(item[0], { x: x + 0.15, y: y + 0.05, w: 4.2, h: 0.32, fontSize: 9, bold: true, color: C.red, fontFace: "Calibri", charSpacing: 1 });
    s.addText(item[1], { x: x + 0.15, y: y + 0.35, w: 4.2, h: 0.6, fontSize: 10.5, color: C.lightGray, fontFace: "Calibri", valign: "top" });
  });
}

// ==============================
// SLIDE 13: TRANSPORT
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  TRANSPORTS — BALVI -> TIRANA", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  // Transport steps visual
  const steps = [
    { label: "BALVI", sub: "Autobuss -> Rīga", detail: "~3 h  -  ~8 €", icon: "BUS" },
    { label: "RĪGA", sub: "Finnair -> Helsinki", detail: "~1h 10min  -  ~80 €", icon: "RIX" },
    { label: "HELSINKI", sub: "Finnair AY -> Tirana", detail: "~3h 20min  -  ~180–280 €", icon: "HEL" },
    { label: "TIRANA", sub: "Galamērķis sasniegts!", detail: "Kopā: ~8 h  -  ~268–368 €", icon: "TIA" },
  ];

  steps.forEach((st, i) => {
    const x = 0.35 + i * 2.35;

    // Circle
    s.addShape(pres.shapes.OVAL, { x: x + 0.6, y: 1.05, w: 1.1, h: 1.1, fill: { color: i === 3 ? C.red : C.gray2 }, line: { color: i === 3 ? C.redLight : C.gray4, width: 2 } });
    s.addText(st.icon, { x: x + 0.6, y: 1.05, w: 1.1, h: 1.1, fontSize: 11, bold: true, color: i === 3 ? C.white : C.red, fontFace: "Calibri", align: "center", valign: "middle" });

    // Arrow between steps
    if (i < 3) {
      s.addShape(pres.shapes.LINE, {
        x: x + 1.75, y: 1.6, w: 0.6, h: 0,
        line: { color: C.red, width: 2 }
      });
    }

    s.addText(st.label, { x, y: 2.25, w: 2.3, h: 0.4, fontSize: 13, bold: true, color: C.white, fontFace: "Calibri", align: "center" });
    s.addText(st.sub, { x, y: 2.62, w: 2.3, h: 0.35, fontSize: 10, color: C.silver, fontFace: "Calibri", align: "center" });
    s.addText(st.detail, { x, y: 2.95, w: 2.3, h: 0.35, fontSize: 10, bold: true, color: C.red, fontFace: "Calibri", align: "center" });
  });

  // Justification box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.5, w: 9.3, h: 1.0, fill: { color: C.gray2 }, line: { color: C.gray3 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.5, w: 0.06, h: 1.0, fill: { color: C.red }, line: { color: C.red } });
  s.addText("GAISA TRANSPORTA IZVĒLES PAMATOJUMS", { x: 0.5, y: 3.55, w: 9.0, h: 0.3, fontSize: 9, bold: true, color: C.red, fontFace: "Calibri", charSpacing: 1 });
  s.addText(
    "Gaisa transports ir vienīgais saprātīgais variants 5 dienu ceļojumam no Latvijas uz Albāniju. " +
    "Citi varianti (autobuss ~36 h, vilciens ar vairākām pārsēšanās) nav praktiski. " +
    "Finnair maršruts caur Helsinkiem — 1 pārsēšanās, stabils savienojums, konkurētspējīgas cenas. " +
    "Avoti: www.finnair.com - google.com/flights",
    { x: 0.5, y: 3.85, w: 9.0, h: 0.6, fontSize: 10.5, color: C.lightGray, fontFace: "Calibri" }
  );

  // Internal transport note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 4.6, w: 9.3, h: 0.75, fill: { color: "1A1A2A" }, line: { color: C.gray4 } });
  s.addText("IEKŠĒJAIS TRANSPORTS ALBĀNIJĀ — nomāts auto vai furgoni (furgon): " +
    "Tirana -> Berat (~2h, ~5€)  -  Berat -> Sarandë (~3.5h, ~8€)  -  Sarandë ↔ Xhamil/Butrint (~30min, ~3€)  -  " +
    "Tirana -> Theth (~5h, ~15€ ar mini-busu)",
    { x: 0.5, y: 4.6, w: 9.0, h: 0.75, fontSize: 10, color: C.silver, fontFace: "Calibri", valign: "middle" }
  );
}

// ==============================
// SLIDE 14: ROUTE
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  CEĻOJUMA MARŠRUTS — 5 DIENAS, 7+ OBJEKTI", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const days = [
    { day: "1. DIENA", city: "TIRANA", places: ["Skanderbega laukums", "Bunk'Art muzejs", "Et'hem Beu mošeja"], color: C.red },
    { day: "2. DIENA", city: "BERAT", places: ["Berat vecpilsēta (UNESCO)", "Berat cietoksnis", "Osmāņu kvartāls"], color: "8B0000" },
    { day: "3. DIENA", city: "XHAMIL", places: ["Xhamil pludmales", "Jonijas jūra", "Sarandë pilsēta"], color: "004080" },
    { day: "4. DIENA", city: "BUTRINT", places: ["Butrint drupas (UNESCO)", "Senā amfiteātra tūre", "Butrint ezers"], color: "005020" },
    { day: "5. DIENA", city: "THETH", places: ["Theth ciems (Alpes)", "Albāņu Alpes taka", "Atgriešanās Tiranā"], color: "404000" },
  ];

  days.forEach((d, i) => {
    const x = 0.35 + i * 1.86;

    // Header
    s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 1.76, h: 0.55, fill: { color: d.color }, line: { color: d.color } });
    s.addText(d.day, { x, y: 0.95, w: 1.76, h: 0.3, fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
    s.addText(d.city, { x, y: 1.2, w: 1.76, h: 0.3, fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });

    // Body
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 1.76, h: 3.85, fill: { color: C.gray2 }, line: { color: C.gray3 } });

    d.places.forEach((p, pi) => {
      const py = 1.6 + pi * 0.62;
      s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: py + 0.07, w: 0.28, h: 0.28, fill: { color: d.color }, line: { color: d.color } });
      s.addText(`${pi + 1}`, { x: x + 0.12, y: py + 0.07, w: 0.28, h: 0.28, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", fontFace: "Calibri" });
      s.addText(p, { x: x + 0.46, y: py, w: 1.22, h: 0.45, fontSize: 9.5, color: C.lightGray, fontFace: "Calibri", valign: "middle" });
    });

    // Transport connector (except last)
    if (i < 4) {
      s.addShape(pres.shapes.LINE, {
        x: x + 1.76, y: 1.22, w: 0.1, h: 0,
        line: { color: C.red, width: 1 }
      });
    }
  });

  // Summary bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 5.1, w: 9.3, h: 0.38, fill: { color: C.red }, line: { color: C.red } });
  s.addText("KOPĀ: 9 apskates objekti  -  5 pilsētas  -  Tirana -> Berat -> Xhamil/Sarandë -> Butrint -> Theth -> Tirana", {
    x: 0.35, y: 5.1, w: 9.3, h: 0.38, fontSize: 10, bold: true, color: C.white,
    fontFace: "Calibri", align: "center", valign: "middle"
  });
}

// ==============================
// SLIDE 15: OBJECT DESCRIPTIONS
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  APSKATES OBJEKTU APRAKSTI", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const objects = [
    ["Skanderbega laukums", "Tirānas sirds — nosaukts nacionālā varoņa vārdā. Ieskauta ar Nacionālo vēstures muzeju un Et'hem Beu mošeju."],
    ["Bunk'Art muzejs", "Komunistu bunkera sistēma, pārveidota par vēstures muzeju. 5 stāvi zem zemes, 1000+ eksponātu par komunismu."],
    ["Berat (UNESCO)", "\"Tūkstoš logu pilsēta\" — viena no skaistākajām viduslaiku pilsētām Balkānos ar osmāņu arhitektūru."],
    ["Berat cietoksnis", "Aktīvi apdzīvots viduslaiku cietoksnis. Unikāls skats uz Osum upi un vecpilsētu."],
    ["Xhamil pludmales", "Jonijas jūras kristāldzidra piekraste — viens no skaistākajiem pludmales galamērķiem Eiropā, nav pārsātināts."],
    ["Butrint (UNESCO)", "Senā pilsēta no 7. gs. pirms Kr. — grieķu, romiešu un bizantiešu kultūru slāņi. Amfiteātris, bazilikas, mūri."],
  ];

  objects.forEach((obj, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.35 + col * 4.8;
    const y = 1.0 + row * 1.52;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.5, h: 1.38,
      fill: { color: C.gray2 }, line: { color: C.gray3 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.38, fill: { color: C.red }, line: { color: C.red } });

    // Number
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.08, w: 0.4, h: 0.4, fill: { color: C.red }, line: { color: C.red } });
    s.addText(`${i + 1}`, { x: x + 0.15, y: y + 0.08, w: 0.4, h: 0.4, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", fontFace: "Calibri" });

    s.addText(obj[0], { x: x + 0.65, y: y + 0.07, w: 3.7, h: 0.38, fontSize: 12, bold: true, color: C.white, fontFace: "Calibri" });
    s.addText(obj[1], { x: x + 0.15, y: y + 0.52, w: 4.2, h: 0.82, fontSize: 10.5, color: C.silver, fontFace: "Calibri", valign: "top" });
  });
}

// ==============================
// SLIDE 16: MAP (Visual)
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  MARŠRUTA KARTE — ALBĀNIJA", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  // Map background
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 0.95, w: 5.8, h: 4.45, fill: { color: "0A1A2A" }, line: { color: C.gray4 } });

  // Simple Albania outline using shapes (approximate)
  // Albania shape - simplified rectangle with key cities plotted
  s.addText("🗺 ALBĀNIJAS KARTE", {
    x: 0.35, y: 0.95, w: 5.8, h: 4.45,
    fontSize: 11, color: "1A3A5A", fontFace: "Calibri", align: "center", valign: "top",
    bold: true
  });

  // Draw approximate map elements
  // Coastline (west side) - blue
  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.2, w: 0.12, h: 3.9, fill: { color: "1A4A8A" }, line: { color: "1A4A8A" } });
  // Country body
  s.addShape(pres.shapes.RECTANGLE, { x: 0.82, y: 1.2, w: 4.4, h: 3.9, fill: { color: "0D2A1A" }, line: { color: "0D2A1A" } });

  // City dots and labels
  const cities = [
    { name: "TIRANA", x: 1.8, y: 2.1, day: "D1", color: C.red },
    { name: "BERAT", x: 2.2, y: 2.8, day: "D2", color: "FF6600" },
    { name: "XHAMIL", x: 1.6, y: 3.8, day: "D3", color: "0088FF" },
    { name: "BUTRINT", x: 1.5, y: 4.3, day: "D4", color: "00AA44" },
    { name: "THETH", x: 2.5, y: 1.6, day: "D5", color: "CCAA00" },
  ];

  // Draw route line
  const cityCoords = cities.map(c => ({ x: c.x + 0.15, y: c.y + 0.15 }));
  for (let i = 0; i < cityCoords.length - 1; i++) {
    const c1 = cityCoords[i];
    const c2 = cityCoords[i + 1];
    const dx = c2.x - c1.x;
    const dy = c2.y - c1.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    s.addShape(pres.shapes.LINE, {
      x: c1.x, y: c1.y, w: dx, h: dy,
      line: { color: C.red, width: 1.5, dashType: "dash" }
    });
  }

  cities.forEach((city) => {
    s.addShape(pres.shapes.OVAL, { x: city.x, y: city.y, w: 0.32, h: 0.32, fill: { color: city.color }, line: { color: C.white, width: 1 } });
    s.addText(city.day, { x: city.x, y: city.y, w: 0.32, h: 0.32, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle", fontFace: "Calibri" });
    s.addText(city.name, { x: city.x + 0.35, y: city.y, w: 1.5, h: 0.32, fontSize: 9.5, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });
  });

  // Adriatic sea label
  s.addText("ADRIJAS\nJŪRA", { x: 0.4, y: 2.5, w: 0.9, h: 0.6, fontSize: 8, color: "4488CC", fontFace: "Calibri", align: "center", italic: true });

  // Legend on right
  const legend = [
    { day: "D1", city: "Tirana", transport: "Lidmašīna (TIA)", color: C.red },
    { day: "D2", city: "Berat", transport: "Furgons ~2h, ~5€", color: "FF6600" },
    { day: "D3", city: "Xhamil/Sarandë", transport: "Autobuss ~3.5h, ~8€", color: "0088FF" },
    { day: "D4", city: "Butrint", transport: "Taksometrs ~30min, ~3€", color: "00AA44" },
    { day: "D5", city: "Theth (Alpes)", transport: "Mini-buss ~5h, ~15€", color: "CCAA00" },
  ];

  s.addText("MARŠRUTS UN IEKŠĒJAIS TRANSPORTS", {
    x: 6.35, y: 0.98, w: 3.3, h: 0.42, fontSize: 9, bold: true, color: C.red,
    fontFace: "Calibri", charSpacing: 1
  });

  legend.forEach((lg, i) => {
    const y = 1.5 + i * 0.78;
    s.addShape(pres.shapes.RECTANGLE, { x: 6.35, y, w: 3.3, h: 0.7, fill: { color: C.gray2 }, line: { color: C.gray3 } });
    s.addShape(pres.shapes.OVAL, { x: 6.45, y: y + 0.19, w: 0.32, h: 0.32, fill: { color: lg.color }, line: { color: lg.color } });
    s.addText(lg.day, { x: 6.45, y: y + 0.19, w: 0.32, h: 0.32, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle", fontFace: "Calibri" });
    s.addText(lg.city, { x: 6.85, y: y + 0.05, w: 2.7, h: 0.3, fontSize: 11, bold: true, color: C.white, fontFace: "Calibri" });
    s.addText(lg.transport, { x: 6.85, y: y + 0.35, w: 2.7, h: 0.28, fontSize: 9.5, color: C.silver, fontFace: "Calibri" });
  });
}

// ==============================
// SLIDE 17: COSTS
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  IZMAKSU KALKULĀCIJA", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const costs = [
    ["Aviobiļetes (Finnair, turp+atpakaļ)", "—", "—", "~260–360 €"],
    ["Autobuss Balvi -> Rīga (turp+atpakaļ)", "—", "—", "~16 €"],
    ["Iekšējais transports Albānijā", "~6–10 €", "4 d.", "~24–40 €"],
    ["Nakšņošana (hostel / 2★ viesnīca)", "~20–35 €", "4 n.", "~80–140 €"],
    ["Ēdināšana (3x dienā)", "~12–20 €", "5 d.", "~60–100 €"],
    ["Ieejas biļetes (Butrint, Bunk'Art u.c.)", "~5–10 €", "5 d.", "~25–50 €"],
    ["Suvenīri un izdevumi", "~5–10 €", "5 d.", "~25–50 €"],
    ["Ceļojumu apdrošināšana", "—", "—", "~25 €"],
  ];

  const colHeaders = ["KATEGORIJA", "€/DIENĀ", "PERIODS", "KOPĀ"];
  const colX = [0.35, 5.7, 7.2, 8.5];
  const colW = [5.2, 1.35, 1.2, 1.4];

  // Headers
  colHeaders.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, { x: colX[ci], y: 1.0, w: colW[ci], h: 0.45, fill: { color: C.red }, line: { color: C.red } });
    s.addText(h, { x: colX[ci], y: 1.0, w: colW[ci], h: 0.45, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
  });

  costs.forEach((row, ri) => {
    const bg = ri % 2 === 0 ? C.gray2 : C.gray1;
    const y = 1.48 + ri * 0.44;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, { x: colX[ci], y, w: colW[ci], h: 0.42, fill: { color: bg }, line: { color: C.gray3 } });
      s.addText(cell, {
        x: colX[ci], y, w: colW[ci], h: 0.42,
        fontSize: ci === 0 ? 10.5 : 11,
        color: ci === 3 ? C.white : C.lightGray,
        bold: ci === 3,
        fontFace: "Calibri", align: ci === 0 ? "left" : "center", valign: "middle",
        margin: ci === 0 ? [0, 0, 0, 8] : 0
      });
    });
  });

  // Total row
  const totalY = 1.48 + costs.length * 0.44;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: totalY, w: 9.3, h: 0.56, fill: { color: C.red }, line: { color: C.red } });
  s.addText("KOPĒJĀS IZMAKSAS", { x: 0.35, y: totalY, w: 5.5, h: 0.56, fontSize: 12, bold: true, color: C.white, fontFace: "Calibri", valign: "middle", margin: [0, 0, 0, 10] });
  s.addText("~50–85 €/dienā", { x: 5.7, y: totalY, w: 2.5, h: 0.56, fontSize: 11, bold: true, color: "FFCCCC", fontFace: "Calibri", align: "center", valign: "middle" });
  s.addText("516–781 €", { x: 8.3, y: totalY, w: 1.6, h: 0.56, fontSize: 13, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
}

// ==============================
// SLIDE 18: RISKS
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("3. POSMS  -  IESPĒJAMIE RISKI UN ANALĪZE", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const risks = [
    { icon: "⚠", title: "CEĻU INFRASTRUKTŪRA", level: "VIDĒJS", color: "CC6600", desc: "Lauku ceļi dažkārt nav asfaltēti (Theth virzienā). Ieteicams SUV vai vietējais mini-buss." },
    { icon: "🦟", title: "VESELĪBA UN KLIMATS", level: "ZEMS", color: "4A8A00", desc: "Vasarā karstuma risks. Piekrastē odi un ērces. Nav obligātu vakcināciju." },
    { icon: "💬", title: "VALODAS BARJERA", level: "ZEMS–VIDĒJS", color: "CC6600", desc: "Lauku rajonos angļu valoda nav pieejama. Tirana un kūrorti — labi. Google Translate ar albāņu val." },
    { icon: "🌊", title: "JŪRAS DROŠĪBA", level: "ZEMS", color: "4A8A00", desc: "Xhamil pludmales reizēm bez glābējiem. Jonijas jūras straumes var būt stipras." },
    { icon: "🏛", title: "POLITISKĀ STABILITĀTE", level: "ĻOTI ZEMS", color: "1A6A1A", desc: "Albānija ir stabila demokrātija un ES kandidātvalsts. Tūrisma drošība ir augstā līmenī." },
    { icon: "💳", title: "FINANSIĀLIE RISKI", level: "VIDĒJS", color: "CC6600", desc: "Daudzas vietas pieņem tikai skaidru naudu (lek – ALL). Bankomatiem var nebūt mazpilsētās." },
  ];

  const levelColors = { "ZEMS": "1A5A1A", "ĻOTI ZEMS": "0A4A0A", "ZEMS–VIDĒJS": "5A4A00", "VIDĒJS": "5A3A00" };

  risks.forEach((r, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.35 + col * 4.85;
    const y = 1.0 + row * 1.52;

    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.55, h: 1.38, fill: { color: C.gray2 }, line: { color: C.gray3 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.38, fill: { color: r.color }, line: { color: r.color } });

    // Level badge
    s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.08, w: 1.5, h: 0.28, fill: { color: levelColors[r.level] || "333300" }, line: { color: "000000", width: 0 } });
    s.addText(r.level, { x: x + 0.15, y: y + 0.08, w: 1.5, h: 0.28, fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });

    s.addText(r.title, { x: x + 1.75, y: y + 0.05, w: 2.65, h: 0.38, fontSize: 9.5, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });
    s.addText(r.desc, { x: x + 0.15, y: y + 0.45, w: 4.3, h: 0.88, fontSize: 10.5, color: C.silver, fontFace: "Calibri", valign: "top" });
  });
}

// ==============================
// SLIDE 19: CONCLUSIONS
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.black };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("SECINĀJUMI UN KOPSAVILKUMS", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  const conclusions = [
    { num: "01", title: "ALBĀNIJA — EIROPAS DĀRGAKMENS", body: "Dienvideiropa - Balkāni - Adrijas un Jonijas jūra\nLētākā Vidusjūras valsts ar unikālu vēsturi un dabu" },
    { num: "02", title: "TRANSPORTS", body: "Rīga -> Helsinki -> Tirana  (Finnair, 1 pārsēšanās)\nKopējais laiks: ~8 h  -  Biļetes: ~268–368 €" },
    { num: "03", title: "MARŠRUTS: 5 DIENAS, 9 OBJEKTI", body: "Tirana -> Berat -> Xhamil -> Butrint -> Theth\n2× UNESCO mantojums - pludmales - kalni" },
    { num: "04", title: "IZMAKSAS", body: "Kopā: ~516–781 €  -  ~50–85 €/dienā\nLētāka par Grieķiju, Horvātiju vai Itāliju" },
    { num: "05", title: "GALVENIE RISKI", body: "Lauku ceļi (risinājums: SUV/mini-buss)\nValodas barjera (risinājums: Google Translate)" },
  ];

  conclusions.forEach((c, i) => {
    const col = i % 2 === 0 && i !== 4 ? i % 2 : 1;
    const x = 0.35 + (i < 4 ? (i % 2) * 4.85 : 2.35);
    const row = Math.floor(i / 2);
    const y = 1.0 + row * 1.52;
    const w = i === 4 ? 5.3 : 4.55;

    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 1.38, fill: { color: C.gray1 }, line: { color: C.gray3 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.38, fill: { color: C.red }, line: { color: C.red } });

    s.addText(c.num, { x: x + 0.15, y: y + 0.06, w: 0.5, h: 0.38, fontSize: 22, bold: true, color: C.red, fontFace: "Arial Black", valign: "middle" });
    s.addText(c.title, { x: x + 0.7, y: y + 0.06, w: w - 0.85, h: 0.38, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", valign: "middle", charSpacing: 1 });
    s.addText(c.body, { x: x + 0.15, y: y + 0.52, w: w - 0.25, h: 0.82, fontSize: 11, color: C.silver, fontFace: "Calibri", valign: "top" });
  });

  // Big recommendation
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 4.85, w: 9.3, h: 0.6, fill: { color: C.red }, line: { color: C.red } });
  s.addText("⭐  IETEIKUMS: DOTIES UZ ALBĀNIJU! — droša, skaista, lēta un vēl nav pārsātināta ar tūristiem.", {
    x: 0.35, y: 4.85, w: 9.3, h: 0.6, fontSize: 13, bold: true, color: C.white,
    fontFace: "Calibri", align: "center", valign: "middle"
  });
}

// ==============================
// SLIDE 20: SOURCES + SELF ASSESSMENT
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.darkGray };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.gray1 }, line: { color: C.gray1 } });
  s.addText("IZMANTOTIE AVOTI  -  PAŠIZVĒRTĒŠANA", {
    x: 0.4, y: 0, w: 9, h: 0.9, fontSize: 13, color: C.red, bold: true,
    fontFace: "Calibri", charSpacing: 2, align: "left", valign: "middle", margin: 0
  });

  // Sources
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 0.98, w: 4.5, h: 0.38, fill: { color: C.red }, line: { color: C.red } });
  s.addText("IZMANTOTIE AVOTI", { x: 0.35, y: 0.98, w: 4.5, h: 0.38, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", charSpacing: 2 });

  const sources = [
    "www.finnair.com — gaisa maršruti un cenas",
    "www.google.com/flights — biļešu cenu salīdzinājums",
    "www.albania.al — Albānijas Tūrisma aģentūra",
    "www.lonelyplanet.com/albania — ceļojumu ceļvedis",
    "whc.unesco.org — UNESCO mantojuma objekti",
    "www.booking.com — nakšņošanas cenu aptauja",
    "www.visit.balvi.lv — Balvu novada tūrisms",
    "Wikipedia — Albānijas ģeogrāfija un vēsture",
    "www.visitventspils.com - www.visitdaugavpils.lv",
  ];

  sources.forEach((src, i) => {
    const y = 1.42 + i * 0.38;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 4.5, h: 0.36, fill: { color: i % 2 === 0 ? C.gray2 : C.gray1 }, line: { color: C.gray3 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.06, h: 0.36, fill: { color: C.red }, line: { color: C.red } });
    s.addText(src, { x: 0.46, y, w: 4.3, h: 0.36, fontSize: 10, color: C.silver, fontFace: "Calibri", valign: "middle" });
  });

  // Self assessment
  s.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 0.98, w: 4.5, h: 0.38, fill: { color: C.red }, line: { color: C.red } });
  s.addText("PAŠIZVĒRTĒŠANA", { x: 5.15, y: 0.98, w: 4.5, h: 0.38, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", charSpacing: 2 });

  const selfItems = [
    ["SASNIEGTS", "Albānija izpētīta visaptveroši — ģeogrāfija, transports, maršruts, izmaksas, riski."],
    ["IEMĀCĪTS", "Reāla ceļojuma plānošana ar budžetu. Albānija — viena no lētākajām Eiropas valstīm."],
    ["DARĪTU SAVĀDĀK", "Veiktu vairāk interviju; pētītu arī Vlorë un dienvidu piekrasti."],
    ["IETEIKUMS", "Personīgā pieredze no Balkāniem ļoti papildinātu projektu."],
  ];

  selfItems.forEach((item, i) => {
    const y = 1.42 + i * 0.86;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.15, y, w: 4.5, h: 0.78, fill: { color: C.gray2 }, line: { color: C.gray3 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.15, y, w: 0.06, h: 0.78, fill: { color: C.red }, line: { color: C.red } });
    s.addText(item[0], { x: 5.25, y: y + 0.05, w: 4.3, h: 0.28, fontSize: 9, bold: true, color: C.red, fontFace: "Calibri", charSpacing: 1 });
    s.addText(item[1], { x: 5.25, y: y + 0.32, w: 4.3, h: 0.42, fontSize: 10.5, color: C.lightGray, fontFace: "Calibri", valign: "top" });
  });
}

// ==============================
// SLIDE 21: THANK YOU
// ==============================
{
  let s = pres.addSlide();
  s.background = { color: C.black };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.red }, line: { color: C.red } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.red }, line: { color: C.red } });

  s.addText("ALBĀNIJA", {
    x: 0, y: 0.5, w: 10, h: 2.5, fontSize: 120, bold: true, color: "141414",
    fontFace: "Arial Black", align: "center", valign: "middle", margin: 0
  });

  s.addText("Paldies par uzmanību!", {
    x: 1, y: 1.4, w: 8, h: 1.2, fontSize: 44, bold: true, color: C.white,
    fontFace: "Arial Black", align: "center", valign: "middle", margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 2.75, w: 3, h: 0.05, fill: { color: C.red }, line: { color: C.red } });

  s.addText("Projekts \"Starptautiskais tūrisms: VIP\"  -  12. klase", {
    x: 1, y: 2.9, w: 8, h: 0.45, fontSize: 14, color: C.silver,
    fontFace: "Calibri", align: "center", margin: 0
  });

  const stages = ["1. POSMS: Viedokļi [OK]", "2. POSMS: Izvērtēšana [OK]", "3. POSMS: Pētījums – Albānija [OK]"];
  stages.forEach((st, i) => {
    const x = 1.0 + i * 2.8;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.5, w: 2.5, h: 0.65, fill: { color: C.gray1 }, line: { color: C.gray3 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.5, w: 2.5, h: 0.06, fill: { color: C.red }, line: { color: C.red } });
    s.addText(st, { x, y: 3.56, w: 2.5, h: 0.59, fontSize: 11, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", bold: true });
  });

  s.addText("Gatavs atbildēt uz jautājumiem!", {
    x: 1, y: 4.4, w: 8, h: 0.5, fontSize: 14, color: C.red, bold: true,
    fontFace: "Calibri", align: "center", italic: true, margin: 0
  });
}

pres.writeFile({ fileName: "./Albania_VIP_Tourism.pptx" })
  .then(() => console.log("Done!"))
  .catch(e => console.error(e));
