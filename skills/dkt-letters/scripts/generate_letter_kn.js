/**
 * Генератор письма в Комитет науки по АЦИВС (Тип 9)
 * АО «Фонд науки» — ДКТ
 *
 * Запуск:
 *   cd /sessions/jolly-magical-wright/letter_work
 *   node generate_letter_kn.js
 *
 * Перед запуском заполнить переменные в секции ДАННЫЕ ПРОЕКТА.
 * Библиотека docx должна быть установлена: npm install docx
 */

const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, BorderStyle } = require('docx');
const fs = require('fs');
const path = require('path');

// ============================================================
// ДАННЫЕ ПРОЕКТА — заполнить перед генерацией
// ============================================================
const DATA = {
  irn: 'DP23692492',
  projectNameRu: 'Инновационная технология производства комбикормов для рыб, птиц и сельскохозяйственных животных',
  projectNameKz: 'Балықтар, құс және ауылшаруашылық малдары үшін комбикорм өндірудің инновациялық технологиясы',
  gpNameRu: 'ТОО «Жамалбек»',
  gpNameKz: '«Жамалбек» ЖШС',
  dogovorNum: '108',
  dogovorDateRu: '24 сентября 2024 года',
  dogovorDateKz: '2024 жылғы 24 қыркүйек',
  nnsDecisionNum: '5',
  nnsDecisionDateRu: '12–18 сентября 2024 года',
  nnsDecisionDateKz: '2024 жылғы 12-18 қыркүйек',
  etap: 2,                    // номер этапа (число)
  acivisDateRu: '03 апреля 2026 года',
  acivisDateKz: '2026 жылғы 03 сәуір',
  ekonomiyaRu: '6 409 616,20 (шесть миллионов четыреста девять тысяч шестьсот шестнадцать) тенге 20 тиын',
  ekonomiyaKz: '6 409 616,20 (алты миллион төрт жүз тоғыз мың алты жүз он алты) теңге 20 тиын',
  escrowAccount: 'KZ918562203141104528',   // '' если экономии нет
  escrowBankRu: 'АО «БанкЦентрКредит»',
  escrowBankKz: '«БанкЦентрКредит» АҚ',
  // Доп. соглашения: [{numRu: '1', dateRu: '19 декабря 2024 года', numKz: '1', dateKz: '2024 жылғы 19 желтоқсан'}, ...]
  ds: [
    { numRu: '1', dateRu: '19 декабря 2024 года',   numKz: '1', dateKz: '2024 жылғы 19 желтоқсан' },
    { numRu: '2', dateRu: '01 декабря 2025 года',    numKz: '2', dateKz: '2025 жылғы 01 желтоқсан' },
  ],
  // Фиксированные (не менять)
  signatoryRu: 'Председатель Правления',
  signatoryNameRu: 'Ашкин А.А.',
  signatoryKz: 'Басқарма төрағасы',
  signatoryNameKz: 'Ашкин А.А.',
  executorKz: 'Орынд.: Снадин А.',
  executorRu: 'Исп.: Снадин А.',
  phone: 'Тел.: 8-7172-57-50-08',
  // Путь к логотипу (извлечь из шаблона заранее)
  logoPath: '/tmp/letter_logo_unpacked/word/media/image1.jpeg',
  // Выходной файл
  outputPath: '/sessions/jolly-magical-wright/mnt/CLO/Письмо_КН_DP23692492_2этап.docx',
};
// ============================================================

const etapOrdRu = ['первого','второго','третьего','четвёртого','пятого'];
const etapOrdKz = ['1-ші','2-ші','3-ші','4-ші','5-ші'];
const etapNumRu = `${DATA.etap}-го`;
const etapNumKz = `${DATA.etap}-кезең`;

function para(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    spacing: { before: opts.before || 0, after: opts.after !== undefined ? opts.after : 100, line: 276, lineRule: 'auto' },
    indent: opts.indent ? { firstLine: 720 } : undefined,
    children: runs.map(r => new TextRun({ text: r.text, font: 'Times New Roman', size: 24, bold: r.bold || false, italics: r.italics || false }))
  });
}
function txt(text, opts = {}) { return para([{ text, ...opts }], opts); }
function empty(before = 0, after = 0) { return new Paragraph({ spacing: { before, after, line: 276 }, children: [new TextRun({ text: '', font: 'Times New Roman', size: 24 })] }); }
function rightBlock(lines) {
  return new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 80, line: 276 },
    children: lines.map((l, i) => new TextRun({ text: l, font: 'Times New Roman', size: 24, bold: true, break: i > 0 ? 1 : 0 })) });
}
function appItem(num, text) {
  return new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 60, line: 276 }, indent: { left: 720 },
    children: [new TextRun({ text: `${num}) `, font: 'Times New Roman', size: 24, italics: true }), new TextRun({ text, font: 'Times New Roman', size: 24, italics: true })] });
}
function divider() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 4 } },
    children: [new TextRun({ text: '', font: 'Times New Roman', size: 8 })] });
}

// Приложения
function buildAppItems(lang) {
  const items = [];
  let n = 1;
  if (lang === 'kz') {
    items.push(appItem(n++, `жобаның ${DATA.etap}-кезеңінің аралық есебі;`));
    items.push(appItem(n++, `${DATA.acivisDateKz}дегі бөлінген қаражаттың мақсатты пайдалануын талдау қорытындысы (АЦИВС);`));
    items.push(appItem(n++, `${DATA.dogovorDateKz}дегі № ${DATA.dogovorNum} ҒҒТҚ коммерцияландыруға грант беру туралы шарт${DATA.ds.length ? ';' : '.'}`));
    DATA.ds.forEach((ds, i) => {
      const last = i === DATA.ds.length - 1;
      items.push(appItem(n++, `${ds.dateKz}дегі № ${ds.numKz} қосымша келісім${last ? '.' : ';'}`));
    });
  } else {
    items.push(appItem(n++, `промежуточный отчет ${DATA.etap}-го этапа Проекта;`));
    items.push(appItem(n++, `заключение АЦИВС от ${DATA.acivisDateRu};`));
    items.push(appItem(n++, `Договор о предоставлении гранта на коммерциализацию РННТД № ${DATA.dogovorNum} от ${DATA.dogovorDateRu}${DATA.ds.length ? ';' : '.'}`));
    DATA.ds.forEach((ds, i) => {
      const last = i === DATA.ds.length - 1;
      items.push(appItem(n++, `Дополнительное соглашение № ${ds.numRu} от ${ds.dateRu}${last ? '.' : ';'}`));
    });
  }
  return items;
}

const logoImage = fs.readFileSync(DATA.logoPath);

// Абзац об экономии (только если есть)
function ekonomiyaParaKz() {
  if (!DATA.escrowAccount) return [];
  return [
    empty(60, 0),
    para([
      { text: 'Жоба бойынша ұсынылған аралық есеп бойынша Күнтізбелік жоспардың барлық іс-шаралары орындалды, жоспарланған сатып алу ' + DATA.etap + '-кезеңнің бөлінген сомасына сәйкес жүзеге асырылды. Сонымен қатар, ' + DATA.etap + '-кезең үшін аралық есеп бойынша бөлінген қаражаттың мақсатты пайдаланылуын талдау қорытындысына сәйкес Жобаны іске асыру барысында ' },
      { text: DATA.ekonomiyaKz, bold: true },
      { text: ' сомасына үнемдеу қалыптасты. Аталған үнемдеу ' + DATA.escrowBankKz + ' эскроу шотында ' + DATA.escrowAccount + ' қалды.' },
    ], { indent: true }),
  ];
}
function ekonomiyaParaRu() {
  if (!DATA.escrowAccount) return [];
  return [
    empty(60, 0),
    para([
      { text: 'Согласно предоставленному промежуточному отчету по Проекту, все мероприятия Календарного плана выполнены, запланированный закуп осуществлен согласно выделенной сумме ' + DATA.etap + '-го этапа. Вместе с тем, согласно заключению анализа целевого использования выделенных средств по промежуточному отчету за ' + DATA.etap + '-й этап, в ходе реализации Проекта образовалась экономия на сумму ' },
      { text: DATA.ekonomiyaRu, bold: true },
      { text: '. Данная экономия остается на эскроу-счёте ' + DATA.escrowBankRu + ' ' + DATA.escrowAccount + '.' },
    ], { indent: true }),
  ];
}

const doc = new Document({
  styles: { default: { document: { run: { font: 'Times New Roman', size: 24 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, right: 720, bottom: 850, left: 1080 } } },
    children: [
      // Логотип + адрес
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 }, children: [new ImageRun({ type: 'jpeg', data: logoImage, transformation: { width: 530, height: 77 }, altText: { title: 'Logo', description: 'Fond', name: 'Logo' } })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [new TextRun({ text: '010000, Астана қаласы, Тәуелсіздік даңғылы 41     010000, город Астана, пр. Тауелсиздик 41', font: 'Times New Roman', size: 18 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 180 }, children: [new TextRun({ text: 'тел +7(7172) 575 003, www.science-fund.kz     тел +7(7172) 575 003, www.science-fund.kz', font: 'Times New Roman', size: 18 })] }),

      // === КАЗАХСКИЙ БЛОК ===
      rightBlock(['Қазақстан Республикасы', 'Ғылым және жоғары білім министрлігінің', 'Ғылым комитеті']),
      empty(60, 60),
      para([
        { text: `Ғылыми және (немесе) ғылыми-техникалық қызметті дамытудың басым бағыты «Ғылыми және (немесе) ғылыми-техникалық қызмет нәтижелерін коммерцияландыру» бойынша Ұлттық ғылыми кеңестің (бұдан әрі – ҰҒК) ${DATA.nnsDecisionDateKz}ғы № ${DATA.nnsDecisionNum} шешімінің негізінде «Ғылым қоры» АҚ (бұдан әрі – Қор) Грант алушы – ${DATA.gpNameKz}-мен (бұдан әрі – Грант алушы) ${DATA.irn} ` },
        { text: `«${DATA.projectNameKz}»`, bold: true },
        { text: ` жобасын (бұдан әрі – Жоба) іске асыру үшін ${DATA.dogovorDateKz}те ғылыми және (немесе) ғылыми-техникалық қызмет нәтижелерін коммерцияландыруға грант беру туралы ` },
        { text: `№ ${DATA.dogovorNum} Шартты`, bold: true },
        { text: ' (бұдан әрі – Шарт) жасасты.' },
      ], { indent: true }),
      empty(60, 0),
      para([
        { text: `Грант алушы Шарттың «Күнтізбелік жоспар» 1-қосымшасына, «Шығыстар сметасы» 2-қосымшасына сәйкес іс-шараларды орындау бойынша, сондай-ақ «Жобаны іске асыру тиімділігін талдау жөніндегі есеп» 4-қосымшасына сәйкес жобаны іске асыру тиімділігіне қол жеткізу бойынша нәтижелері бойынша ` },
        { text: `${DATA.etap}-кезеңнің аралық есебін берді.`, bold: true },
      ], { indent: true }),
      ...ekonomiyaParaKz(),
      empty(60, 0),
      txt('Қазақстан Республикасы Ғылым және жоғары білім министрінің «Ұлттық ғылыми кеңестер туралы ережені бекіту туралы» 2023 жылғы 25 қыркүйектегі № 487 бұйрығының 3-тарауы 20-тармағының 8) тармақшасына сәйкес Ұлттық ғылыми кеңестің (бұдан әрі – ҰҒК) негізгі міндеттерінің бірі ғылыми және (немесе) ғылыми-техникалық қызмет, ғылыми және (немесе) ғылыми-техникалық қызмет нәтижелерін коммерцияландыру туралы аралық және қорытынды есептерді, сондай-ақ сараптама орталығы ұсынатын ғылыми, ғылыми-техникалық жобалар мен бағдарламалардың, ғылыми және (немесе) ғылыми-техникалық қызмет нәтижелерін коммерцияландыру жобаларының іске асырылу мониторингінің қорытындыларын қарау, осындай есептер мен мониторингтің қорытындыларын қараудың нәтижелері бойынша шешім қабылдау болып табылады.', { indent: true }),
      empty(60, 0),
      para([
        { text: `Жоғарыда айтылғандарды ескере отырып, Қор № ${DATA.irn} ` },
        { text: `«${DATA.projectNameKz}»`, bold: true },
        { text: ` жобасы бойынша ` },
        { text: `${DATA.etap}-кезеңнің аралық есебін`, bold: true },
        { text: ' және ' },
        { text: `${DATA.etap}-кезең үшін аралық есеп бойынша бөлінген қаражаттың мақсатты пайдаланылуына талдау қорытындысын`, bold: true },
        { text: ' «Ғылыми және (немесе) ғылыми-техникалық қызмет нәтижелерін коммерцияландыру» ғылыми және (немесе) ғылыми-техникалық қызметті дамытудың басым бағыты бойынша ҰҒК қарауына Жобаны одан әрі қаржыландыру жөнінде шешім қабылдау үшін шығаруды сұрайды.' },
      ], { indent: true }),
      empty(80, 0),
      txt('Қосымша:', { italics: true, after: 60 }),
      ...buildAppItems('kz'),
      empty(100, 0),
      txt(`${DATA.signatoryKz}                                              ${DATA.signatoryNameKz}`, { bold: true }),
      empty(80, 0),
      txt(DATA.executorKz, { italics: true, after: 40 }),
      txt(DATA.phone, { italics: true, after: 200 }),

      divider(),
      empty(60, 0),

      // === РУССКИЙ БЛОК ===
      rightBlock(['Комитет науки', 'Министерства науки и высшего', 'образования Республики Казахстан']),
      empty(60, 60),
      para([
        { text: `На основании решения Национального научного совета (далее – ННС) от ${DATA.nnsDecisionDateRu} № ${DATA.nnsDecisionNum} по приоритетному направлению развития научной и (или) научно-технической деятельности «Коммерциализация результатов научной и (или) научно-технической деятельности» АО «Фонд науки» (далее – Фонд) заключил Договор о предоставлении гранта на коммерциализацию результатов научной и (или) научно-технической деятельности ` },
        { text: `№ ${DATA.dogovorNum} от ${DATA.dogovorDateRu}`, bold: true },
        { text: ` (далее – Договор) по проекту № ${DATA.irn} ` },
        { text: `«${DATA.projectNameRu}»`, bold: true },
        { text: ' (далее – Проект) с Грантополучателем – ' },
        { text: DATA.gpNameRu, bold: true },
        { text: '.' },
      ], { indent: true }),
      empty(60, 0),
      para([
        { text: 'Грантополучателем предоставлен ' },
        { text: `промежуточный отчет за ${DATA.etap}-й этап`, bold: true },
        { text: ' по исполнению мероприятий согласно Приложению 1 «Календарный план», Приложению 2 «Смета расходов», а также результатов по достижению эффективности реализации проекта Приложения 4 «Отчет по анализу эффективности реализации проекта» Договора.' },
      ], { indent: true }),
      ...ekonomiyaParaRu(),
      empty(60, 0),
      txt('В соответствии с пп. 8) п. 20 главы 3 Приказа Министра науки и высшего образования Республики Казахстан от 25 сентября 2023 года № 487 «Об утверждении положения о национальных научных советах» одним из основных задач Национального научного совета (далее – ННС) является рассмотрение промежуточных и итоговых отчетов о научной и (или) научно-технической деятельности, коммерциализации результатов научной и (или) научно-технической деятельности, а также итогов мониторинга реализации научных, научно-технических проектов и программ, проектов коммерциализации результатов научной и (или) научно-технической деятельности, представляемых центром экспертизы, принятие решения по результатам рассмотрения таких отчетов и итогов мониторинга для вынесения заключения.', { indent: true }),
      empty(60, 0),
      para([
        { text: 'Учитывая вышеизложенное, Фонд просит внести на рассмотрение ННС по приоритетному направлению развития научной и (или) научно-технической деятельности «Коммерциализация результатов научной и (или) научно-технической деятельности» ' },
        { text: `промежуточный отчет ${DATA.etap}-го этапа`, bold: true },
        { text: ` по проекту № ${DATA.irn} ` },
        { text: `«${DATA.projectNameRu}»`, bold: true },
        { text: ` и заключение анализа целевого использования выделенных средств по промежуточному отчету за ${DATA.etap}-й этап по Проекту для принятия решения по дальнейшему финансированию проекта.` },
      ], { indent: true }),
      empty(80, 0),
      txt('Приложение:', { italics: true, after: 60 }),
      ...buildAppItems('ru'),
      empty(100, 0),
      txt(`${DATA.signatoryRu}                                       ${DATA.signatoryNameRu}`, { bold: true }),
      empty(80, 0),
      txt(DATA.executorRu, { italics: true, after: 40 }),
      txt(DATA.phone, { italics: true }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(DATA.outputPath, buf);
  console.log('✅ Файл сохранён:', DATA.outputPath);
}).catch(err => { console.error('❌ Ошибка:', err); process.exit(1); });
