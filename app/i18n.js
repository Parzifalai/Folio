// Язык интерфейса. Тексты в коде написаны по-русски; при другом языке модуль сам
// подменяет их на экране (кнопки, подсказки, меню, уведомления). Содержимое документов не трогаем.
const I18N = {
  // русский: [английский, казахский]
  'Общий': ['General', 'Жалпы'],
  'Число': ['Number', 'Сан'],
  'Целое': ['Whole number', 'Бүтін сан'],
  'Деньги, ₸': ['Currency, ₸', 'Ақша, ₸'],
  'Процент': ['Percent', 'Пайыз'],
  'Дата': ['Date', 'Күн'],
  'Дата и время': ['Date and time', 'Күн және уақыт'],
  'Адрес ячейки — введите, чтобы перейти': ['Cell address — type to jump', 'Ұяшық мекенжайы — өту үшін енгізіңіз'],
  'Новый лист': ['New sheet', 'Жаңа парақ'],
  'Сумма': ['Sum', 'Сомасы'],
  'Среднее': ['Average', 'Орташа'],
  'Количество': ['Count', 'Саны'],
  'Ширина': ['Width', 'Ені'],
  'Высота': ['Height', 'Биіктігі'],
  'Такой ячейки нет — пример: B12': ['No such cell — example: B12', 'Мұндай ұяшық жоқ — мысалы: B12'],
  'В формуле ошибка — проверьте скобки и кавычки': ['The formula has an error — check brackets and quotes', 'Формулада қате бар — жақшалар мен тырнақшаларды тексеріңіз'],
  'Эту формулу Folio пока не считает — Excel посчитает её при открытии': ['Folio can’t calculate this formula yet — Excel will when the file is opened', 'Folio бұл формуланы әзірге есептемейді — Excel ашқанда есептейді'],
  'Слишком большая область': ['The area is too large', 'Аймақ тым үлкен'],
  'Переименовать': ['Rename', 'Атын өзгерту'],
  'Удалить лист': ['Delete sheet', 'Парақты жою'],
  'Лист': ['Sheet', 'Парақ'],
  'Данные на нём пропадут.': ['Its data will be lost.', 'Ондағы деректер жоғалады.'],
  'Вырезать': ['Cut', 'Қию'],
  'Копировать': ['Copy', 'Көшіру'],
  'Вставить строку выше': ['Insert row above', 'Жоғарыға жол қосу'],
  'Вставить строку ниже': ['Insert row below', 'Төменге жол қосу'],
  'Вставить столбец слева': ['Insert column left', 'Сол жаққа баған қосу'],
  'Вставить столбец справа': ['Insert column right', 'Оң жаққа баған қосу'],
  'Удалить строки': ['Delete rows', 'Жолдарды жою'],
  'Удалить столбцы': ['Delete columns', 'Бағандарды жою'],
  'Ширина по содержимому': ['Fit width to content', 'Мазмұнға сай ені'],
  'Нажмите Ctrl+V': ['Press Ctrl+V', 'Ctrl+V басыңыз'],
  'Отменить': ['Undo', 'Болдырмау'],
  'Вернуть': ['Redo', 'Қайтару'],
  'Заливка': ['Fill', 'Бояу'],
  'Как в Excel (числа справа)': ['Like Excel (numbers right)', 'Excel сияқты (сандар оңда)'],
  'Переносить по словам': ['Wrap text', 'Сөздер бойынша тасымалдау'],
  'Формат чисел': ['Number format', 'Сан пішімі'],
  'Автосумма': ['AutoSum', 'Автосома'],
  'Строки и столбцы': ['Rows and columns', 'Жолдар мен бағандар'],
  'Без заливки': ['No fill', 'Бояусыз'],
  'Таблица Excel': ['Excel spreadsheet', 'Excel кестесі'],
  'Диаграмма': ['Chart', 'Диаграмма'],
  'Объект': ['Object', 'Нысан'],
  'Заголовок слайда': ['Slide title', 'Слайд тақырыбы'],
  'Текст слайда': ['Slide text', 'Слайд мәтіні'],
  'Новый слайд': ['New slide', 'Жаңа слайд'],
  'Дублировать слайд': ['Duplicate slide', 'Слайдты көшіру'],
  'Удалить слайд': ['Delete slide', 'Слайдты жою'],
  'Переместить выше': ['Move up', 'Жоғары жылжыту'],
  'Переместить ниже': ['Move down', 'Төмен жылжыту'],
  'Показывать слайд': ['Show slide', 'Слайдты көрсету'],
  'Скрыть слайд при показе': ['Hide slide in slideshow', 'Көрсетілімде слайдты жасыру'],
  'Перемещение слайда': ['moving a slide', 'слайдты жылжыту'],
  'В презентации должен остаться хотя бы один слайд': ['A presentation needs at least one slide', 'Презентацияда кемінде бір слайд болуы керек'],
  'Удаление слайда': ['deleting a slide', 'слайдты жою'],
  'Копия слайда': ['duplicating a slide', 'слайд көшірмесі'],
  'Скрытие слайда': ['hiding a slide', 'слайдты жасыру'],
  'Новый слайд ': ['new slide', 'жаңа слайд'],
  'Презентация': ['Presentation', 'Презентация'],
  'Слайды': ['Slides', 'Слайдтар'],
  'Показ с этого слайда': ['Slideshow from this slide', 'Осы слайдтан көрсету'],
  'Титульный слайд': ['Title Slide', 'Титул слайды'],
  'Заголовок и объект': ['Title and Content', 'Тақырып және нысан'],
  'Заголовок раздела': ['Section Header', 'Бөлім тақырыбы'],
  'Два объекта': ['Two Content', 'Екі нысан'],
  'Сравнение': ['Comparison', 'Салыстыру'],
  'Только заголовок': ['Title Only', 'Тек тақырып'],
  'Пустой слайд': ['Blank', 'Бос слайд'],
  'Объект с подписью': ['Content with Caption', 'Жазуы бар нысан'],
  'Рисунок с подписью': ['Picture with Caption', 'Жазуы бар сурет'],
  'Новая таблица Excel': ['New Excel spreadsheet', 'Жаңа Excel кестесі'],
  'Новая презентация': ['New presentation', 'Жаңа презентация'],
  'Таблица.xlsx': ['Spreadsheet.xlsx', 'Кесте.xlsx'],
  'Презентация.pptx': ['Presentation.pptx', 'Презентация.pptx'],
  'Презентация PowerPoint': ['PowerPoint presentation', 'PowerPoint презентациясы'],
  'Открываются PDF, Word, Excel, PowerPoint и картинки': ['Folio opens PDF, Word, Excel, PowerPoint and images', 'PDF, Word, Excel, PowerPoint және суреттер ашылады'],
  'Сохранить': ['Save', 'Сақтау'],
  'Сохранить как…': ['Save as…', 'Басқаша сақтау…'],
  'Показать в папке': ['Show in folder', 'Қалтада көрсету'],
  'Перенести в левую половину': ['Move to the left half', 'Сол жақ жартысына ауыстыру'],
  'Перенести в правую половину': ['Move to the right half', 'Оң жақ жартысына ауыстыру'],
  'Открыть рядом': ['Open side by side', 'Қатар ашу'],
  'Закрыть': ['Close', 'Жабу'],
  'Открыть файл…': ['Open file…', 'Файлды ашу…'],
  'Новый документ Word': ['New Word document', 'Жаңа Word құжаты'],
  'Пустая половина': ['Empty half', 'Бос жарты'],
  'Если не сохранить, правки пропадут.': ['If you don’t save, your changes will be lost.', 'Сақтамасаңыз, өзгерістер жоғалады.'],
  'Не сохранять': ['Don’t save', 'Сақтамау'],
  'Отмена': ['Cancel', 'Болдырмау'],
  'Сохранено как новый файл': ['Saved as a new file', 'Жаңа файл ретінде сақталды'],
  'Сохранено': ['Saved', 'Сақталды'],
  'Файл открыт в другой программе — закройте её или сохраните как новый файл': ['The file is open in another program — close it or save as a new file', 'Файл басқа бағдарламада ашық — оны жабыңыз немесе жаңа файл ретінде сақтаңыз'],
  'Файл защищён паролем — такие пока не открываются': ['The file is password-protected — such files can’t be opened yet', 'Файл құпиясөзбен қорғалған — мұндай файлдар әзірге ашылмайды'],
  'Файл не найден — возможно, его переместили': ['File not found — it may have been moved', 'Файл табылмады — мүмкін, ол жылжытылған'],
  'Не получилось открыть файл': ['Couldn’t open the file', 'Файлды ашу мүмкін болмады'],
  'Картинки.pdf': ['Images.pdf', 'Суреттер.pdf'],
  'Новый документ.pdf': ['New document.pdf', 'Жаңа құжат.pdf'],
  'Документ.docx': ['Document.docx', 'Құжат.docx'],
  'Готовлю показ…': ['Preparing the show…', 'Көрсетілім дайындалуда…'],
  'Не получилось начать показ': ['Couldn’t start the show', 'Көрсетілімді бастау мүмкін болмады'],
  'Полный экран': ['Full screen', 'Толық экран'],
  'выйти — Esc или F11': ['exit — Esc or F11', 'шығу — Esc немесе F11'],
  'Новый PDF': ['New PDF', 'Жаңа PDF'],
  'Открыть…': ['Open…', 'Ашу…'],
  'Показ на весь экран': ['Full-screen show', 'Толық экранда көрсету'],
  'Один документ': ['One document', 'Бір құжат'],
  'Два документа рядом': ['Two documents side by side', 'Екі құжат қатар'],
  'Показать панель инструментов': ['Show toolbar', 'Құралдар тақтасын көрсету'],
  'Скрыть панель инструментов': ['Hide toolbar', 'Құралдар тақтасын жасыру'],
  'Выйти из полного экрана': ['Exit full screen', 'Толық экраннан шығу'],
  'Во весь экран': ['Full screen', 'Толық экран'],
  'Пока открываются PDF, Word и картинки': ['Only PDF, Word and images can be opened for now', 'Әзірге тек PDF, Word және суреттер ашылады'],
  'Перетащите к любому краю': ['Drag to any edge', 'Кез келген шетке сүйреңіз'],
  'двойной щелчок — вниз': ['double-click — back to the bottom', 'екі рет басу — төменге'],
  'размер — за край панели': ['resize by dragging an edge', 'өлшемі — панель шетінен тартыңыз'],
  'Обычный': ['Normal', 'Қалыпты'],
  'Название': ['Title', 'Атауы'],
  'Подзаголовок': ['Subtitle', 'Тақырыпша'],
  'Абзац списка': ['List paragraph', 'Тізім абзацы'],
  'Подпись к рисунку': ['Caption', 'Сурет жазуы'],
  'Без интервала': ['No spacing', 'Аралықсыз'],
  'Цитата': ['Quote', 'Дәйексөз'],
  'Заголовок оглавления': ['TOC heading', 'Мазмұн тақырыбы'],
  'Документ Word': ['Word document', 'Word құжаты'],
  'Документ': ['Document', 'Құжат'],
  'Документ.pdf': ['Document.pdf', 'Құжат.pdf'],
  'Сохранено в PDF': ['Saved as PDF', 'PDF ретінде сақталды'],
  'Добавить строку выше': ['Insert row above', 'Жоғарыға жол қосу'],
  'Добавить строку ниже': ['Insert row below', 'Төменге жол қосу'],
  'Удалить строку': ['Delete row', 'Жолды жою'],
  'Удалить таблицу': ['Delete table', 'Кестені жою'],
  'Страницы': ['Pages', 'Беттер'],
  'Стиль абзаца': ['Paragraph style', 'Абзац стилі'],
  'Подчёркнутый': ['Underline', 'Асты сызылған'],
  'Цвет текста': ['Text color', 'Мәтін түсі'],
  'Выделение цветом': ['Highlight', 'Түспен белгілеу'],
  'Выравнивание': ['Alignment', 'Туралау'],
  'По левому краю': ['Align left', 'Сол жақ шет бойынша'],
  'По центру': ['Center', 'Ортасы бойынша'],
  'По правому краю': ['Align right', 'Оң жақ шет бойынша'],
  'По ширине': ['Justify', 'Ені бойынша'],
  'Маркированный список': ['Bulleted list', 'Таңбаланған тізім'],
  'Нумерованный список': ['Numbered list', 'Нөмірленген тізім'],
  'Картинка': ['Image', 'Сурет'],
  'Таблица': ['Table', 'Кесте'],
  'Без выделения': ['No highlight', 'Белгілеусіз'],
  'Добавлена пустая страница': ['Blank page added', 'Бос бет қосылды'],
  'Открыть или создать': ['Open or create', 'Ашу немесе жасау'],
  'Новая вкладка': ['New tab', 'Жаңа қойынды'],
  'Главная': ['Home', 'Басты бет'],
  'Ещё': ['More', 'Тағы'],
  'Откройте второй документ': ['Open a second document', 'Екінші құжатты ашыңыз'],
  'Перетащите файл сюда или нажмите, чтобы открыть': ['Drop a file here or click to open', 'Файлды осында сүйреңіз немесе ашу үшін басыңыз'],
  'Новый документ': ['New document', 'Жаңа құжат'],
  'Перетащите панель влево или вправо': ['Drag the panel left or right', 'Панельді солға немесе оңға сүйреңіз'],
  'Открываю…': ['Opening…', 'Ашылуда…'],
  'Выбор': ['Select', 'Таңдау'],
  'Текст': ['Text', 'Мәтін'],
  'T — щёлкните по тексту, чтобы исправить, или по пустому месту, чтобы добавить': ['T — click text to edit it, or an empty spot to add new text', 'T — түзету үшін мәтінді, жаңасын қосу үшін бос жерді басыңыз'],
  'Перо': ['Pen', 'Қалам'],
  'Маркер': ['Highlighter', 'Маркер'],
  'Фигуры': ['Shapes', 'Пішіндер'],
  'Замазать': ['Cover up', 'Бояп жасыру'],
  'Подпись': ['Signature', 'Қолтаңба'],
  'Скрыть панель': ['Hide toolbar', 'Тақтаны жасыру'],
  'Показать панель': ['Show toolbar', 'Тақтаны көрсету'],
  'Шрифт': ['Font', 'Қаріп'],
  'Размер': ['Size', 'Өлшем'],
  'Жирный': ['Bold', 'Қалың'],
  'Курсив': ['Italic', 'Көлбеу'],
  'Прямоугольник': ['Rectangle', 'Тіктөртбұрыш'],
  'Овал': ['Ellipse', 'Сопақ'],
  'Линия': ['Line', 'Сызық'],
  'Стрелка': ['Arrow', 'Көрсеткі'],
  'Толщина': ['Thickness', 'Қалыңдығы'],
  'Дублировать': ['Duplicate', 'Көшірмесін жасау'],
  'Удалить': ['Delete', 'Жою'],
  'Картинки': ['Images', 'Суреттер'],
  'Вставить эту подпись': ['Insert this signature', 'Осы қолтаңбаны қою'],
  'Очистить': ['Clear', 'Тазалау'],
  'Вставить': ['Insert', 'Қою'],
  'Вставить пустую страницу': ['Insert a blank page', 'Бос бет қосу'],
  'Повернуть': ['Rotate', 'Бұру'],
  'Удалить страницу': ['Delete page', 'Бетті жою'],
  'Добавить страницы из файла': ['Add pages from a file', 'Файлдан беттер қосу'],
  'Повернуть вправо': ['Rotate right', 'Оңға бұру'],
  'Повернуть влево': ['Rotate left', 'Солға бұру'],
  'Сохранить страницу отдельным PDF': ['Save page as a separate PDF', 'Бетті бөлек PDF ретінде сақтау'],
  'Переместить': ['Move', 'Жылжыту'],
  'Копировать сюда': ['Copy here', 'Осында көшіру'],
  'Страница скопирована': ['Page copied', 'Бет көшірілді'],
  'Нельзя удалить все страницы': ['You can’t delete all pages', 'Барлық беттерді жоюға болмайды'],
  'Добавлена 1 страница': ['1 page added', '1 бет қосылды'],
  'Сохранено отдельным файлом': ['Saved as a separate file', 'Бөлек файл ретінде сақталды'],
  'Найти в документе': ['Find in document', 'Құжаттан іздеу'],
  'Предыдущее': ['Previous', 'Алдыңғы'],
  'Следующее': ['Next', 'Келесі'],
  'Выйти': ['Exit', 'Шығу'],
  '← → листать': ['← → to flip', '← → парақтау'],
  'Esc — выход': ['Esc — exit', 'Esc — шығу'],
  'Конец показа — щёлкните или нажмите любую клавишу, чтобы выйти': ['End of show — click or press any key to exit', 'Көрсетілім аяқталды — шығу үшін басыңыз немесе кез келген пернені басыңыз'],
  'Тема': ['Theme', 'Көрініс'],
  'Как в системе': ['Same as system', 'Жүйедегідей'],
  'Светлая': ['Light', 'Ашық'],
  'Тёмная': ['Dark', 'Қараңғы'],
  'Язык': ['Language', 'Тіл'],
};

// Фразы с числами и именами.
const I18N_RULES = [
  [/^Удалить лист «(.+)»\?$/, ['Delete sheet “$1”?', '«$1» парағын жою керек пе?']],
  [/^Отменено: (.*)$/, ['Undone: $1', 'Болдырылмады: $1']],
  [/^Таблица (\d+)\.xlsx$/, ['Spreadsheet $1.xlsx', 'Кесте $1.xlsx']],
  [/^Презентация (\d+)\.pptx$/, ['Presentation $1.pptx', 'Презентация $1.pptx']],
  [/^Сохранить изменения в «(.+)»\?$/, ['Save changes to “$1”?', '«$1» өзгерістерін сақтау керек пе?']],
  [/^Сохранить (\d+) стр\. отдельным PDF$/, ['Save $1 pages as a separate PDF', '$1 бетті бөлек PDF ретінде сақтау']],
  [/^Удалить (\d+) стр\.$/, ['Delete $1 pages', '$1 бетті жою']],
  [/^Скопировано страниц: (\d+)$/, ['Pages copied: $1', 'Көшірілген беттер: $1']],
  [/^Добавлено страниц: (\d+)$/, ['Pages added: $1', 'Қосылған беттер: $1']],
  [/^Не получилось сохранить: (.*)$/, ['Couldn’t save: $1', 'Сақтау мүмкін болмады: $1']],
  [/^Заголовок (\d)$/, ['Heading $1', 'Тақырып $1']],
  [/^Оглавление (\d)$/, ['TOC $1', 'Мазмұн $1']],
  [/^Новый документ (\d+)\.pdf$/, ['New document $1.pdf', 'Жаңа құжат $1.pdf']],
  [/^Документ (\d+)\.docx$/, ['Document $1.docx', 'Құжат $1.docx']],
];

let LANG = (window.folio && window.folio.prefs && window.folio.prefs.lang) || 'ru';

function T(s) {
  if (LANG === 'ru' || typeof s !== 'string' || !s) return s;
  const k = LANG === 'en' ? 0 : 1;
  if (I18N[s]) return I18N[s][k];
  for (const [re, out] of I18N_RULES) if (re.test(s)) return s.replace(re, out[k]);
  // «Действие · Ctrl+X» — переводим каждую часть
  if (s.includes(' · ')) {
    const parts = s.split(' · ');
    const tr = parts.map((p) => T(p));
    if (tr.some((p, i) => p !== parts[i])) return tr.join(' · ');
  }
  return s;
}

// Где текст — это документ, а не интерфейс.
const I18N_SKIP = '.wd-surface, .page, .t-paper, .wd-thumb-page, .show-stage, .r-dir, .r-name, .tab-name, .xl-tabs, .xl-grid, .xl-bar input, .pp-stage, .pp-thumbs, .pp-notes';
const i18nText = new WeakMap(); // текстовый узел → { orig, tr }
const i18nAttr = new WeakMap(); // элемент → { 'data-tip': {orig, tr}, placeholder: {...} }

function i18nNode(n) {
  if (n.nodeType === 3) {
    const pe = n.parentElement;
    if (!pe || pe.closest(I18N_SKIP)) return;
    const cur = n.data;
    if (!/[А-Яа-яЁё]|[A-Za-z]/.test(cur)) return;
    let rec = i18nText.get(n);
    if (!rec || (cur !== rec.tr && cur !== rec.orig)) rec = { orig: cur.trim() === cur ? cur : cur };
    const want = T(rec.orig.trim());
    rec.tr = want === rec.orig.trim() ? rec.orig : rec.orig.replace(rec.orig.trim(), want);
    i18nText.set(n, rec);
    if (n.data !== rec.tr) n.data = rec.tr;
    return;
  }
  if (n.nodeType !== 1 || n.closest(I18N_SKIP)) return;
  for (const a of ['data-tip', 'placeholder']) {
    if (!n.hasAttribute(a)) continue;
    const cur = n.getAttribute(a);
    const map = i18nAttr.get(n) || {};
    let rec = map[a];
    if (!rec || (cur !== rec.tr && cur !== rec.orig)) rec = { orig: cur };
    rec.tr = T(rec.orig);
    map[a] = rec;
    i18nAttr.set(n, map);
    if (cur !== rec.tr) n.setAttribute(a, rec.tr);
  }
  if (n.matches && n.matches(I18N_SKIP)) return;
  for (const c of n.childNodes) i18nNode(c);
}

const i18nObserver = new MutationObserver((recs) => {
  for (const r of recs) {
    if (r.type === 'characterData') i18nNode(r.target);
    else if (r.type === 'attributes') i18nNode(r.target);
    else for (const a of r.addedNodes) i18nNode(a);
  }
});

function setLang(lang) {
  LANG = lang;
  document.documentElement.lang = lang === 'kk' ? 'kk' : lang;
  i18nNode(document.body);
}

document.addEventListener('DOMContentLoaded', () => {
  i18nObserver.observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['data-tip', 'placeholder'] });
  setLang(LANG);
});
