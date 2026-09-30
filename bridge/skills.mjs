/**
 * Роутер скиллов АЦИВС.
 *
 * Читает skills/<имя>/SKILL.md, вытаскивает ключевые фразы из description
 * (после «упоминает:» / «Триггеры:») и по тексту команды выбирает скилл.
 * Выбранный скилл дописывается к запросу как служебная подсказка — агент
 * открывает SKILL.md и выполняет его инструкции без лишних вопросов.
 *
 * Явные команды тоже работают: «скилл acivs-memo», «/acivs-memo», «навигатор».
 */

import { readdirSync, readFileSync, existsSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

export const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
export const SKILLS_DIR = join(ROOT, 'skills')

/** Нормализация: нижний регистр, ё→е, лишние знаки — в пробел. */
const norm = (s) =>
  s
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"'“”.,;:!?()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

/** Грубая основа: первые 5 букв длинных слов («проверь» = «проверить», «записки» = «записка»). */
const stem = (w) => (w.length > 5 ? w.slice(0, 5) : w)
const FUNCTION_WORDS = new Set('на в во и по для с со к о об от до из за при про или а но не что как'.split(' '))
/** Значимые основы фразы, без служебных слов. */
const tokens = (p) =>
  norm(p)
    .split(' ')
    .filter((w) => w && !FUNCTION_WORDS.has(w))
    .map(stem)

function parseSkill(dir) {
  const file = join(SKILLS_DIR, dir, 'SKILL.md')
  if (!existsSync(file)) return null
  const raw = readFileSync(file, 'utf8')
  const fm = /^---\n([\s\S]*?)\n---/.exec(raw)?.[1] ?? ''
  const name = (/^name:\s*(.+)$/m.exec(fm)?.[1]?.trim() ?? dir).replace(/^["']|["']$/g, '')
  const descBlock = /^description:\s*>?\s*\n?([\s\S]*)$/m.exec(fm)?.[1] ?? ''
  const description = descBlock.replace(/\n\s+/g, ' ').replace(/^["']|["']$/g, '').trim()

  // Фразы-триггеры: всё между «упоминает:»/«Триггеры:» и следующим предложением.
  const m = /(?:упоминает|триггеры|упоминаются)[^:]*:\s*([^]*?)(?:\.\s+[А-ЯA-Z]|$)/i.exec(description)
  const phrases = (m?.[1] ?? '')
    .split(/[,;]/)
    .map((p) => p.trim())
    .filter((p) => p.length >= 4 && p.length <= 60)
  // Запасной вариант: если триггеров в description нет — берём заголовок скилла.
  if (!phrases.length) phrases.push(description.split(/[.:]/)[0].slice(0, 80))

  // Добавляем имя скилла и его «человеческие» варианты.
  phrases.push(name.replace(/^"|"$/g, ''), name.replace(/^"|"$/g, '').replace(/-/g, ' '))
  return {
    name,
    dir,
    description,
    phrases: [...new Set(phrases)].map((p) => ({ raw: p, tok: tokens(p) })).filter((p) => p.tok.length),
  }
}

let cache = null
export function loadSkills() {
  if (cache) return cache
  try {
    cache = readdirSync(SKILLS_DIR, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => parseSkill(d.name))
      .filter(Boolean)
  } catch {
    cache = []
  }
  return cache
}

/** Ручные псевдонимы для голосовых команд — то, что говорят, а не то, что в description. */
const ALIASES = {
  'acivs-navigator': ['навигатор', 'что делать с отчетом', 'ациввс помоги', 'помоги с ацивс'],
  'dogovor-navigator': ['навигатор договора', 'заключение договора'],
  'acivs-memo': ['сз на транш', 'сз на аванс', 'служебка на транш', 'служебка на аванс'],
  'escrow-fot': ['расписка на зарплату', 'эскроу на зарплату', 'расписка на фот'],
  'acivs-red-flags': ['красные флаги', 'риск скор', 'проверь на нарушения'],
  'spravka-status-proekta': ['справка по проекту', 'статус проекта'],
  'portfolio-dashboard': ['дашборд портфеля', 'покажи портфель'],
  'perevodchik-gylym-kory': ['переведи на казахский', 'аудар', 'на казахский'],
  'gp-mailing': ['рассылка гп', 'разошли гп'],
  'kontragent-check': ['проверь контрагента', 'проверка контрагента', 'проверь бин', 'аффилированность поставщика'],
  'acivs-knowledge': ['какой лимит', 'что говорит кд', 'правила конкурсной документации', 'можно ли по кд', 'как считать штраф'],
  'acivs-budget-check': ['проверь бюджет', 'лимиты по смете', 'проверь фот'],
  'dogovor-draft-check': ['проверь проект договора', 'проверь договор гранта'],
  'dogovor-fill': ['заполни договор', 'сверь договор с эталоном'],
  'acivs-conclusion': ['заключение ацивс', 'собери заключение'],
  'acivs-docs-check': ['чек-лист документов', 'комплектность документов'],
  'acivs-activities': ['проверь мероприятия', 'план против факта'],
  'acivs-financial': ['финансовая часть ацивс', 'раздел 2 и 3'],
  'acivs-intake': ['разбери пакет документов', 'классифицируй документы'],
  'dkt-letters-acivs-notice': ['уведомление по ацивс', 'хабарлама'],
  'investkom': ['инвестком', 'пз на инвестком'],
  'karta-vnedreniya': ['карта внедрения'],
  'perechislenie-tracker': ['статус перечислений', 'трекер перечислений'],
  'sz-fed-perechislenie': ['сз в фэд', 'служебка в фэд на перечисление'],
  'nns-letters-table': ['реестр писем на ннс'],
  'audit-kd-rnntd': ['аудит конкурсной документации'],
}

/**
 * @returns {{ skill: object, hits: string[], score: number }[]} лучшие совпадения, по убыванию.
 */
export function matchSkills(text, limit = 2) {
  const said = norm(text)
  if (!said) return []
  const saidSet = new Set(said.split(' ').map(stem))
  const skills = loadSkills()

  // Явное «скилл X» или «/X».
  const explicit = /(?:скилл|навык|skill|\/)\s*([a-z0-9-]+)/.exec(said)
  if (explicit) {
    const s = skills.find((k) => k.name === explicit[1])
    if (s) return [{ skill: s, hits: [explicit[1]], score: 100 }]
  }

  const scored = []
  for (const skill of skills) {
    const hits = []
    let score = 0
    for (const p of skill.phrases) {
      // Все значимые слова фразы есть в команде (порядок не важен).
      // Одиночное слово короче 5 букв слишком шумное — пропускаем.
      const single = p.tok.length === 1
      if (single && p.tok[0].length < 5) continue
      if (p.tok.every((t) => saidSet.has(t))) {
        hits.push(p.raw)
        score += single ? 1 : 1 + p.tok.length // длинные фразы весомее
      }
    }
    for (const a of ALIASES[skill.name] ?? []) {
      if (tokens(a).every((t) => saidSet.has(t))) {
        hits.push(a)
        score += 4
      }
    }
    if (score > 0) scored.push({ skill, hits, score })
  }
  scored.sort((a, b) => b.score - a.score)
  return scored.slice(0, limit)
}

/**
 * Оборачивает команду пользователя подсказкой для агента.
 * Возвращает исходный текст, если ни один скилл не подошёл.
 */
export function routeCommand(text) {
  const found = matchSkills(text)
  if (!found.length) return { text, matched: [] }
  const lines = found.map(
    (f) =>
      `- ${f.skill.name} (совпало: ${f.hits.slice(0, 4).join(', ')}) → skills/${f.skill.dir}/SKILL.md`,
  )
  const hint =
    `\n\n[АВТОМАРШРУТ — служебная подсказка, вслух не читать]\n` +
    `По ключевым словам подобраны скиллы АЦИВС:\n${lines.join('\n')}\n` +
    `Прочитай SKILL.md первого скилла (и нужные файлы из его references/) и выполни задачу по его инструкции. ` +
    `Если не хватает входных файлов — один короткий вопрос, где они лежат. ` +
    `Готовые документы сохраняй в папку output/ и назови файл вслух коротко.`
  return { text: text + hint, matched: found.map((f) => f.skill.name) }
}
