/**
 * Шаблонный скрипт генерации Пояснительной записки для ИнвестКома.
 * Адаптируй переменные в блоке DATA под конкретный проект.
 * Запуск: node create_pz.js
 * Зависимость: npm install docx (если не установлен)
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');
const path = require('path');

// =============================================
// DATA — заполни под конкретный проект
// =============================================
const DATA = {
  // Реквизиты
  gp: 'СПК «MELON Juice Co»',           // Наименование ГП
  irn: 'DP23692015',                     // ИРН проекта
  projectName: 'Комплексная переработка плодов дыни', // Название проекта
  dogovorNum: '69',                      // Номер договора
  dogovorDate: '29.08.2024',             // Дата договора

  // Письмо ГП
  letterNum: '18.03/01',                 // Исх. номер письма ГП
  letterDate: '18.03.2026',              // Дата письма

  // Тип операции: 'включение' или 'замена'
  operationType: 'включение',

  // Кандидат
  candidateFull: 'Спандияр Несибели Турарбеккызы', // ФИО полностью
  candidateShort: 'Спандияр Н.Т.',      // ФИО сокращённо
  candidateDOB: '08.04.1994',            // Дата рождения

  // Должность
  position: 'менеджера',                // В родительном падеже (для текста)
  positionNom: 'менеджер',              // В именительном (для штатного)
  stavka: '1,0',                        // Ставка
  salary: '175 000',                    // Оклад в тенге

  // Образование кандидата (массив объектов)
  education: [
    { year: '2017', degree: 'магистр', university: 'Казахский национальный аграрный университет', specialty: 'Транспорт, транспортная техника и технологии' },
    { year: '2015', degree: 'бакалавр', university: 'Казахский национальный университет имени аль-Фараби', specialty: 'Нанотехнология в электронике' },
  ],

  // Опыт (краткое описание для ПЗ)
  experienceSummary: 'более 8 лет в сфере управления, делопроизводства и администрирования',
  experienceItems: [
    'Март 2024 г. – настоящее время: NOMAD insurance – страховой агент;',
    'Август–Декабрь 2024 г.: FasTracKids – администратор;',
    'Сентябрь 2017 г. – Март 2024 г.: АО «Номад Иншуранас» – специалист по страхованию (6 лет 7 месяцев).',
  ],

  // Функциональные обязанности
  functions: [
    'организация и ведение административной работы в рамках проекта;',
    'подготовка, оформление и контроль внутренней документации (приказы, служебные записки, отчёты);',
    'координация взаимодействия между участниками проектной группы и структурными подразделениями;',
    'контроль сроков исполнения задач и ведение деловой переписки;',
    'подготовка отчётности по административной части проекта и обеспечение документооборота.',
  ],

  // Выходной файл
  outputPath: './ПЗ_ИнвестКом_DP23692015_Спандияр.docx',
};
// =============================================

const font = 'Times New Roman';
const sz = 24; // 12pt

function run(text, opts = {}) {
  return new TextRun({ text, font, size: sz, bold: opts.bold || false, italics: opts.italics || false, underline: opts.underline });
}

function p(children, opts = {}) {
  return new Paragraph({
    children,
    alignment: opts.alignment || AlignmentType.BOTH,
    indent: opts.indent,
    spacing: { line: 276, lineRule: 'auto', ...(opts.spacing || {}) },
    ...opts,
  });
}

function centerP(children) {
  return p(children, { alignment: AlignmentType.CENTER });
}

function indentP(children) {
  return p(children, { indent: { firstLine: 709 } });
}

function emptyLine() {
  return p([run('')]);
}

function bulletItem(text) {
  return new Paragraph({
    children: [run(text)],
    numbering: { reference: 'dashes', level: 0 },
    alignment: AlignmentType.BOTH,
    spacing: { line: 276, lineRule: 'auto' },
  });
}

// Строим тему ПЗ
const topicVerb = DATA.operationType === 'замена'
  ? `О замене ${DATA.positionNom}а по проекту`
  : `О включении участника в проектную группу по проекту`;

// Строим финансовый абзац
const financePara = DATA.operationType === 'замена'
  ? `Замена членов команды по Договору не несёт никаких финансовых затрат. Внесение изменений в Договор не требуется.`
  : `Включение ${DATA.candidateShort} в состав проектной группы на вакантную должность ${DATA.positionNom}а не несёт никаких дополнительных финансовых затрат. Внесение изменений в Договор не требуется, поскольку должность ${DATA.positionNom}а предусмотрена действующим штатным расписанием по Проекту (ставка – ${DATA.stavka}; оклад – ${DATA.salary} тенге).`;

// Строим резолюцию
const resolutionText = DATA.operationType === 'замена'
  ? `Учитывая вышеизложенное, вопрос «О замене участников основного состава и изменении штатного расписания по Договору о предоставлении гранта на коммерциализацию результатов научной и (или) научно-технической деятельности №${DATA.dogovorNum} от ${DATA.dogovorDate} г.» выносится на рассмотрение Инвестиционного комитета.`
  : `Учитывая вышеизложенное, вопрос «О включении ${DATA.candidateFull} в состав проектной группы и штатное расписание по Договору о предоставлении гранта на коммерциализацию результатов научной и (или) научно-технической деятельности №${DATA.dogovorNum} от ${DATA.dogovorDate} г.» выносится на рассмотрение Инвестиционного комитета.`;

const doc = new Document({
  numbering: {
    config: [{
      reference: 'dashes',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '-',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 }, spacing: { line: 276, lineRule: 'auto' } } },
      }],
    }],
  },
  styles: {
    default: {
      document: { run: { font, size: sz }, paragraph: { spacing: { line: 276, lineRule: 'auto' } } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1418, right: 851, bottom: 1134, left: 1701 },
      },
    },
    children: [
      // Заголовок
      centerP([run('ПОЯСНИТЕЛЬНАЯ ЗАПИСКА', { bold: true })]),
      centerP([run('по вопросу повестки дня заседания')]),
      centerP([run('Инвестиционного комитета АО «Фонд науки»')]),
      p([run(`«${topicVerb} ${DATA.irn} «${DATA.projectName}»»`)], { indent: { left: 651 } }),
      emptyLine(),

      // Абзац 1: ГП и Договор
      indentP([
        run(`${DATA.gp} (далее – Грантополучатель) реализует проект ${DATA.irn} «${DATA.projectName}» (далее – Проект), согласно Договору о предоставлении гранта на коммерциализацию результатов научной и (или) научно-технической деятельности №${DATA.dogovorNum} от ${DATA.dogovorDate} г. (далее – Договор).`)
      ]),

      // Абзац 2: Письмо ГП
      indentP([
        run('Грантополучатель обратился письмом в АО «Фонд науки» исх. №'),
        run(DATA.letterNum),
        run(` от ${DATA.letterDate} г. касательно включения в состав проектной группы `),
        run(DATA.candidateFull, { bold: true }),
        run(' на вакантную должность '),
        run(DATA.positionNom + 'а', { bold: true, underline: { type: 'single' } }),
        run(` с занятостью на полную (${DATA.stavka}) ставку.`),
      ]),

      // Функции
      indentP([run(`Функциональные обязанности ${DATA.positionNom}а в Проекте следующие:`)]),
      ...DATA.functions.map(f => bulletItem(f)),
      emptyLine(),

      // Сведения о кандидате
      indentP([run('Сведения о кандидате:')]),
      indentP([run(`${DATA.candidateFull}, дата рождения: ${DATA.candidateDOB} г.`)]),
      indentP([run('Образование:')]),
      ...DATA.education.map(e => bulletItem(`${e.year} г. – ${e.degree}, ${e.university}, специальность «${e.specialty}».`)),
      emptyLine(),
      indentP([run(`Опыт работы: ${DATA.experienceSummary}. В том числе:`)]),
      ...DATA.experienceItems.map(e => bulletItem(e)),
      emptyLine(),

      // Финансовый блок
      indentP([run(financePara)]),
      emptyLine(),

      // Резолюция
      indentP([run(resolutionText)]),
      emptyLine(),

      // Приложение
      p([run('Приложение:', { italics: true })]),
      p([run(`1. Копия письма исх. №${DATA.letterNum} от ${DATA.letterDate} г. и пакет сопутствующих документов.`, { italics: true })]),
      emptyLine(),
      emptyLine(),

      // Подпись
      new Paragraph({
        children: [
          run('Директор ДКТ'),
          new TextRun({ text: '\t', font }),
          run('С. Мажикенов', { bold: true }),
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        alignment: AlignmentType.LEFT,
        spacing: { line: 276, lineRule: 'auto' },
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(DATA.outputPath, buffer);
  console.log(`✅ Файл создан: ${path.resolve(DATA.outputPath)}`);
}).catch(err => console.error('❌ Ошибка:', err));
