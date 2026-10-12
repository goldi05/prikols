# -*- coding: utf-8 -*-
"""Генерация отчета по курсовой работе «Кинодневник» в стиле KursovayaRabotaTerentyev.docx.

Документ-шаблон открывается, его структура (титульный лист, задание, содержание,
введение/заключение, список литературы) переиспользуется, а содержательные разделы
(1..9) полностью переписываются под фактический код репозитория (MVVM, SQLite,
WPF, CinemaDiary).
"""
import copy
import re

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

SRC = 'KursovayaRabotaTerentyev.docx'
DST = 'Otchet_Kursovaya_Kinodnevnik.docx'

CODE_FONT = 'Courier New'
CODE_SIZE = 10.0

d = docx.Document(SRC)
body = d.element.body


def find_para(substr, style=None, start=0):
    for i, p in enumerate(d.paragraphs):
        if i < start:
            continue
        if substr in p.text and (style is None or p.style.name == style):
            return i
    raise ValueError('paragraph not found: ' + substr)


def find_table_after(pidx):
    el = d.paragraphs[pidx]._p.getnext()
    while el is not None:
        if el.tag == qn('w:tbl'):
            for i, t in enumerate(d.tables):
                if t._tbl is el:
                    return i
        el = el.getnext()
    raise ValueError('table not found after paragraph ' + str(pidx))


def remove_between(start_el, stop_el):
    """Удалить элементы между XML-элементами start_el и stop_el (не включая их)."""
    cur = start_el.getnext()
    while cur is not None and cur is not stop_el:
        nxt = cur.getnext()
        cur.getparent().remove(cur)
        cur = nxt


def add_par(text='', style=None, align='justify', indent=True, mono=False,
            bold=False, size=None, page_break=False):
    p = d.add_paragraph()
    if style:
        try:
            p.style = d.styles[style]
        except KeyError:
            pass
    if align == 'justify':
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if indent:
        p.paragraph_format.first_line_indent = 709
    r = p.add_run(text)
    if mono:
        r.font.name = CODE_FONT
        r._element.rPr.rFonts.set(qn('w:eastAsia'), CODE_FONT)
        r.font.size = Pt(size or CODE_SIZE)
    else:
        if bold:
            r.bold = True
        if size:
            r.font.size = Pt(size)
    if page_break:
        p.paragraph_format.page_break_before = True
    return p


def heading_par(text, num=None, page_break=True, center=False):
    p = d.add_paragraph(style='Мой заголовок')
    if page_break:
        p.paragraph_format.page_break_before = True
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = 0
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    full = (f'{num}\t{text}' if num is not None else text)
    r = p.add_run(full)
    r.bold = True
    return p


def code_block(lines):
    for line in lines.rstrip('\n').split('\n'):
        add_par(line, indent=False, mono=True)


def make_toc_field(p):
    def run(el):
        r = p.add_run()._r
        r.append(el)
        return r
    fld = p.add_run()._r
    b = fld.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'begin'})
    b.set(qn('w:dirty'), 'true')
    fld.append(b)
    instr = p.add_run()._r
    it = instr.makeelement(qn('w:instrText'), {})
    it.set(qn('xml:space'), 'preserve')
    it.text = ' TOC \\o "1-3" \\h \\z \\u '
    instr.append(it)
    sep = p.add_run()._r
    sep.append(sep.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'separate'}))
    txt = p.add_run('Содержание будет обновлено при открытии документа (F9).')
    txt.runs[0].font.color.rgb = None
    end = p.add_run()._r
    end.append(end.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'end'}))


# ----------------------------------------------------------------------------
# 1. Индексы опорных абзацев и таблиц (до удаления старого содержания)
# ----------------------------------------------------------------------------
i_h1 = find_para('Постановка задачи', style='Мой заголовок')
i_h2 = find_para('Анализ предметной области', style='Мой заголовок')
i_h3 = find_para('Используемые методы решения', style='Мой заголовок')
i_h4 = find_para('Описание разработанной базы данных', style='Мой заголовок')
i_h5 = find_para('Диаграммы, описывающие внутреннюю структуру')
i_h6 = find_para('Описание классов приложения')
i_h7 = find_para('Описание разработанного приложения с точки зрения пользователя')
i_h8 = find_para('Код приложения', style='Мой заголовок')
i_zakl = find_para('ЗАКЛЮЧЕНИЕ', style='Мой заголовок')
i_lit = find_para('СПИСОК ЛИТЕРАТУРЫ', style='Мой заголовок')

# XML-элементы опорных абзацев (индексы больше не используются после этого блока)
el_h1 = d.paragraphs[i_h1]._p
el_h2 = d.paragraphs[i_h2]._p
el_h3 = d.paragraphs[i_h3]._p
el_h4 = d.paragraphs[i_h4]._p
el_h5 = d.paragraphs[i_h5]._p
el_h6 = d.paragraphs[i_h6]._p
el_h7 = d.paragraphs[i_h7]._p
el_h8 = d.paragraphs[i_h8]._p
el_zakl = d.paragraphs[i_zakl]._p
el_lit = d.paragraphs[i_lit]._p
el_toc = d.paragraphs[find_para('СОДЕРЖАНИЕ', style='Стиль1')]._p
el_intro = d.paragraphs[find_para('Курсовой проект - РАБОТА')]._p

tbl_ents_old = find_table_after(i_h2)      # Таблица 2.1 – Сущности домашней фильмотеки
tbl_abs_old = find_table_after(i_h4)       # Таблица 4.1 – Список основных абстракций
tbl_classes_old = find_table_after(i_h6)   # Таблица 6.1 – Описание классов
tbl_film_old = find_table_after(find_para('Таблица 6.2'))
tbl_filter_old = find_table_after(find_para('Таблица 6.3'))
tbl_repo_old = find_table_after(find_para('Таблица 6.4'))
tbl_cmd_old = find_table_after(find_para('Таблица 6.5'))
tbl_mwvm_old = find_table_after(find_para('Таблица 6.6'))
tbl_media_old = find_table_after(find_para('Таблица 6.7'))

old_content_tables = [tbl_ents_old, tbl_abs_old, tbl_classes_old, tbl_film_old,
                      tbl_filter_old, tbl_repo_old, tbl_cmd_old, tbl_mwvm_old,
                      tbl_media_old]

# ----------------------------------------------------------------------------
# 2. Обновление титульного листа и задания (работа с runs, без смены индексов)
# ----------------------------------------------------------------------------
def set_par_text(el_p, text):
    runs = el_p.findall(qn('w:r'))
    for r in runs[1:]:
        el_p.remove(r)
    if runs:
        ts = list(runs[0].iter(qn('w:t')))
        for t in ts[1:]:
            t.getparent().remove(t)
        if ts:
            ts[0].text = text

set_par_text(d.paragraphs[6]._p, 'Разработка приложения «Кинодневник»')

el_theme = d.paragraphs[find_para('Тема: Разработка приложения')]._p
set_par_text(el_theme, 'Тема: Разработка приложения «Кинодневник»')

el_zad_intro = d.paragraphs[find_para('Курсовая работа посвящена разработке десктопного')]._p
set_par_text(el_zad_intro,
             'Курсовая работа посвящена разработке десктопного Windows-приложения «Кинодневник» '
             'с использованием технологии WPF на языке C# (.NET 8). Приложение предназначено для '
             'ведения личного дневника просмотренных фильмов. Разрабатываемое приложение должно '
             'включать в себя следующую функциональность:')

el_zad_body = d.paragraphs[find_para('Необходимо разработать приложение для ведения домашней фильмотеки')]._p
set_par_text(el_zad_body,
             'Необходимо разработать приложение для ведения домашнего кинодневника. Данная программа '
             'позволяет хранить сведения о фильмах коллекции, их жанрах и режиссёрах, сохранять оценку '
             'и отзыв о просмотренном фильме, а также накапливать статистику в локальной базе данных SQLite.')

# ----------------------------------------------------------------------------
# 3. Удаление старых разделов 1–9 (содержимое между заголовками)
# ----------------------------------------------------------------------------
remove_between(el_h1, el_h2)
remove_between(el_h2, el_h3)
remove_between(el_h3, el_h4)
remove_between(el_h4, el_h5)
remove_between(el_h5, el_h6)
remove_between(el_h6, el_h7)
remove_between(el_h7, el_h8)
remove_between(el_h8, el_zakl)

# ----------------------------------------------------------------------------
# 4. Вспомогательные функции для вставки после указанного элемента
# ----------------------------------------------------------------------------
def insert_par_after(anchor_el, text='', style=None, align='justify', indent=True,
                     mono=False, bold=False, size=None, keep_with_next=False):
    p = d.add_paragraph()
    if style:
        try:
            p.style = d.styles[style]
        except KeyError:
            pass
    if align == 'justify':
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if indent:
        p.paragraph_format.first_line_indent = 709
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    if mono:
        r.font.name = CODE_FONT
        r._element.rPr.rFonts.set(qn('w:eastAsia'), CODE_FONT)
        r.font.size = Pt(size or CODE_SIZE)
    else:
        if bold:
            r.bold = True
        if size:
            r.font.size = Pt(size)
    anchor_el.addnext(p._p)
    return p._p


def insert_heading_after(anchor_el, num, text):
    p = insert_par_after(anchor_el, '', style='Мой заголовок', align='justify',
                         indent=False, keep_with_next=True)
    p.paragraph_format.page_break_before = True
    r = p.makeelement(qn('w:r'), {})
    rp = p.makeelement(qn('w:rPr'), {})
    b = rp.makeelement(qn('w:b'), {})
    rp.append(b)
    r.append(rp)
    t = p.makeelement(qn('w:t'), {})
    t.text = f'{num}\t{text}'
    r.append(t)
    p.append(r)
    return p


def insert_code_after(anchor_el, code):
    last = anchor_el
    for line in code.strip('\n').split('\n'):
        last = insert_par_after(last, line, indent=False, mono=True,
                                keep_with_next=True)
    return last


def insert_caption_after(anchor_el, text):
    return insert_par_after(anchor_el, text, align='center', bold=True)


def fill_table(tbl, rows):
    while len(tbl.rows) < len(rows):
        tbl.add_row()
    excess = []
    for ri in range(len(rows), len(tbl.rows)):
        excess.append(tbl.rows[ri]._tr)
    for tr in excess:
        tr.getparent().remove(tr)
    for row_vals, row in zip(rows, tbl.rows):
        while len(row.cells) < len(row_vals):
            row._tr.append(row._tr.makeelement(qn('w:tc'), {}))
        for cell, val in zip(row.cells, row_vals):
            cell.text = val
            for para in cell.paragraphs:
                for rr in para.runs:
                    rr.font.name = 'Times New Roman'
                    rr.font.size = Pt(12)


def insert_table_after(anchor_el, rows, widths=None):
    idx = len(d.tables)
    new_tbl = d.add_table(rows=len(rows), cols=len(rows[0]))
    new_tbl.style = d.tables[0].style
    el = new_tbl._tbl
    el.getparent().remove(el)
    anchor_el.addnext(el)
    fill_table(d.tables[idx], rows)
    return el


# ----------------------------------------------------------------------------
# РАЗДЕЛ 1. ПОСТАНОВКА ЗАДАЧИ
# ----------------------------------------------------------------------------
a = el_h1
a = insert_par_after(a, 'Тема курсового проекта: «Кинодневник».')
a = insert_par_after(a, 'Постановка задачи:')
for t in ['Описать предметную область ведения личного дневника просмотренных фильмов.',
          'Разработать приложение, которое позволит хранить информацию о фильмах, жанрах и режиссёрах, '
          'добавлять фильмы в коллекцию, удалять их, а также сохранять оценку и отзыв о просмотренном фильме.',
          'Реализовать хранение статистических параметров (общее количество фильмов, количество просмотренных '
          'фильмов, средняя оценка, дата последнего обновления) в базе данных.',
          'Обеспечить выполнение CRUD-операций (create, read, update, delete) над данными базы данных.',
          'Реализовать автоматическое заполнение базы данных тестовыми данными при первом запуске.',
          'Подготовить документацию, описывающую назначение приложения, его архитектуру, основные функции, '
          'порядок запуска и использования.']:
    a = insert_par_after(a, t, style='List Paragraph')
a = insert_par_after(a, 'В базе данных должны храниться сведения об имеющихся в коллекции фильмах. '
                        'Для каждого фильма необходимо учитывать название, год выхода, описание, жанр и режиссёра. '
                        'Для просмотренных фильмов дополнительно хранятся индивидуальная оценка, текст отзыва и дата просмотра. '
                        'Жанры и режиссёры вынесены в отдельные справочники, что исключает дублирование одинаковых записей.')
a = insert_par_after(a, 'Приложение должно обеспечивать просмотр списка всех фильмов, выбор фильма из списка, '
                        'просмотр подробной информации о выбранном фильме, добавление нового фильма по форме, '
                        'удаление выбранного фильма, сохранение оценки и отзыва, а также управление тестовыми данными '
                        'и отображение накопленной статистики.')
a = insert_par_after(a, 'Кроме того, приложение должно обеспечивать удобный пользовательский интерфейс, построенный '
                        'по архитектурному подходу MVVM: навигационное меню слева и переключаемая область содержимого, '
                        'в которой отображается активный раздел (главное меню, список фильмов, добавление фильма, '
                        'детали и отзыв, статистика).')
a = insert_par_after(a, 'В результате выполнения курсового проекта ожидается получение функционального десктопного '
                        'приложения, предназначенного для ведения домашнего кинотека с локальной базой данных SQLite.')

# ----------------------------------------------------------------------------
# РАЗДЕЛ 2. АНАЛИЗ ПРЕДМЕТНОЙ ОБЛАСТИ
# ----------------------------------------------------------------------------
a = el_h2
a = insert_par_after(a, 'Теперь проанализируем предметную область.')
a = insert_par_after(a, 'Домашняя фильмотека в форме приложения «Кинодневник» представляет собой информационную '
                        'систему, предназначенную для хранения, систематизации и просмотра сведений о фильмах, '
                        'которые пользователь смотрел или планирует посмотреть. Объектом учёта выступает фильм, '
                        'который характеризуется названием, годом выхода, описанием, жанром и режиссёром. Если фильм '
                        'просмотрен, то с ним связывается отзыв: оценка, комментарий и дата просмотра.')
a = insert_par_after(a, 'Предметная область курсового проекта включает в себя процессы добавления фильмов в коллекцию, '
                        'их просмотра в списке, редактирования (сохранения отзыва и оценки) и удаления, а также '
                        'накопления статистики о коллекции. Жанры и режиссёры являются справочными сущностями: '
                        'одна и та же запись может относиться к нескольким фильмам, поэтому они хранятся отдельно, '
                        'а связь с ними организована через внешние ключи.')
a = insert_par_after(a, 'После запуска приложения пользователь получает доступ к главному меню. Далее он может '
                        'перейти в раздел «Фильмы» и увидеть список всех записей коллекции, выбрать интересующий фильм, '
                        'перейти в раздел «Детали и отзыв» для просмотра карточки фильма и сохранения своей оценки. '
                        'В разделе «Добавить фильм» создаётся новая запись, а в разделе «Статистика» отображаются '
                        'накопленные показатели и выполняется генерация либо очистка тестовых данных.')
a = insert_par_after(a, 'Выделим сущности и их атрибуты. Сущности приложения «Кинодневник» и их свойства '
                        'представлены в таблице 2.1.')
a = insert_caption_after(a, 'Таблица 2.1 – Сущности предметной области приложения «Кинодневник»')
a = insert_table_after(a, [
    ['Название сущности', 'Атрибуты'],
    ['Фильм', 'Название, год выхода, описание, признак тестовых данных, ссылка на режиссёра, ссылка на жанр'],
    ['Режиссер', 'ФИО режиссёра'],
    ['Жанр', 'Название жанра'],
    ['Отзыв', 'Оценка, комментарий, дата просмотра, ссылка на фильм'],
    ['Статистика', 'Количество фильмов, количество просмотренных фильмов, средняя оценка, дата обновления'],
])
a = insert_par_after(a, 'Таким образом, основными сущностями предметной области являются фильм, режиссёр, жанр, '
                        'отзыв и статистика. Фильм связан с одним режиссёром и одним жанром, а отзыв связан с '
                        'конкретным фильмом и существует только для просмотренных картин.')

# ----------------------------------------------------------------------------
# РАЗДЕЛ 3. ИСПОЛЬЗУЕМЫЕ МЕТОДЫ РЕШЕНИЯ
# ----------------------------------------------------------------------------
methods = [
    ('Разделение программы на отдельные слои.',
     'При разработке приложения использовано разделение решения на три части: Model, ViewModel и View. '
     'Слой Model содержит классы предметной области и репозиторий, слой ViewModel — модели представления и '
     'навигацию, слой View — окно и пользовательские контролы WPF.'),
    ('Архитектурный подход MVVM.',
     'Интерфейс приложения описан средствами XAML и связан с моделями представления через Data Binding. '
     'ViewModel не содержат ссылок на элементы управления, что позволяет независимо изменять интерфейс и логику.'),
    ('Командный подход к обработке действий пользователя.',
     'Для выполнения действий из интерфейса используется класс RelayCommand. С его помощью кнопки окна связываются '
     'с методами моделей представления, а свойство CanExecute автоматически отключает недоступные команды '
     '(например, кнопку удаления при невыбранном фильме).'),
    ('Централизованное хранилище данных (Data Store).',
     'Для согласованной работы всех разделов реализован класс MoviesDataStore. Он один раз загружает данные из '
     'репозитория, хранит выбранный фильм, предоставляет коллекции для связанных списков и уведомляет модели '
     'представления об изменениях через механизм PropertyChanged.'),
    ('Реляционная модель хранения данных.',
     'Для хранения информации была использована база данных SQLite. В приложении реализованы таблицы для фильмов, '
     'жанров, режиссёров, отзывов и статистики, настроены внешние ключи и каскадное удаление отзывов при удалении фильма.'),
    ('Использование репозитория для работы с базой данных.',
     'Для взаимодействия с базой данных реализован класс ConcreteRepository, реализующий интерфейс IRepository. '
     'В нём сосредоточены методы чтения, сохранения и удаления данных, а также методы инициализации структуры БД.'),
    ('Загрузка и инициализация данных при первом запуске.',
     'В проекте реализована автоматическая инициализация базы данных. При первом запуске программа создаёт файл '
     'базы данных, формирует таблицы и заполняет их тестовыми данными. Предусмотрена миграция устаревшей структуры '
     'таблицы Films с текстовым полем Genre.'),
    ('Навигация через флаг видимости представлений.',
     'Переключение разделов реализовано классом ViewModelNavigator и интерфейсом INavigableViewModel: при переходе '
     'у активного представления устанавливается IsVisible = true, у остальных — false, а XAML-разметка преобразует '
     'это значение в видимость соответствующего пользовательского контрола.'),
    ('Сохранение статистических параметров в базе данных.',
     'После каждого изменения данных приложение пересчитывает статистику (количество фильмов, количество '
     'просмотренных фильмов, средняя оценка) и сохраняет её в отдельной таблице Statistics в режиме upsert.'),
]
a = el_h3
a = insert_par_after(a, 'В проекте «Разработка приложения “Кинодневник”» были использованы следующие методы решения:')
for title, desc in methods:
    a = insert_par_after(a, title, style='List Paragraph', bold=True)
    a = insert_par_after(a, desc, style='List Paragraph')
a = insert_par_after(a, 'В целом использование разделения по слоям, подхода MVVM, командного подхода, централизованного '
                        'хранилища данных, реляционной базы данных, репозитория, автоматической инициализации и '
                        'накопления статистики позволило создать простое по архитектуре и удобное в использовании '
                        'приложение, которое легко расширяется новыми разделами и полями.')

# ----------------------------------------------------------------------------
# РАЗДЕЛ 4. ОПИСАНИЕ БАЗЫ ДАННЫХ
# ----------------------------------------------------------------------------
a = el_h4
a = insert_par_after(a, 'Для выбранной предметной области выделим основные абстракции, на основе которых, в дальнейшем '
                        'будет построена концептуальная модель. Перечень абстракций приведён в таблице 4.1.')
a = insert_caption_after(a, 'Таблица 4.1 – Список основных абстракций')
a = insert_table_after(a, [
    ['Имя сущности', 'Атрибуты'],
    ['Фильм', 'ID фильма, название фильма, год выхода, описание, ID режиссёра, ID жанра, признак тестовых данных'],
    ['Режиссёр', 'ID режиссёра, ФИО'],
    ['Жанр', 'ID жанра, название жанра'],
    ['Отзыв', 'ID отзыва, ID фильма, оценка, комментарий, дата просмотра'],
    ['Статистика', 'ID записи, количество фильмов, количество просмотренных фильмов, средняя оценка, дата обновления'],
])
a = insert_par_after(a, 'На основе анализа предметной области построим концептуальную модель. Данная модель '
                        'представлена на рисунке 4.1.')
a = insert_caption_after(a, 'Рисунок 4.1 – Концептуальная модель в виде ER – диаграммы')
a = insert_par_after(a, 'На основе построенной ER – диаграммы разработаем реляционную модель базы данных (рисунок 4.2), '
                        'в которой выделим первичные и внешние ключи. Структура таблиц приведена в таблицах 4.2–4.6.')
a = insert_caption_after(a, 'Таблица 4.2 – Реляционная модель БД')
a = insert_table_after(a, [
    ['Таблица', 'Поле', 'Тип', 'Ограничения'],
    ['Movies', 'Id', 'INTEGER', 'PRIMARY KEY AUTOINCREMENT'],
    ['Movies', 'Title', 'TEXT', 'NOT NULL'],
    ['Movies', 'ReleaseYear', 'INTEGER', 'NOT NULL'],
    ['Movies', 'Description', 'TEXT', ''],
    ['Movies', 'DirectorId', 'INTEGER', 'FOREIGN KEY REFERENCES Directors(Id)'],
    ['Movies', 'GenreId', 'INTEGER', 'FOREIGN KEY REFERENCES Genres(Id)'],
    ['Movies', 'IsTestData', 'INTEGER', 'NOT NULL DEFAULT 0'],
    ['Directors', 'Id', 'INTEGER', 'PRIMARY KEY AUTOINCREMENT'],
    ['Directors', 'FullName', 'TEXT', 'NOT NULL UNIQUE'],
    ['Genres', 'Id', 'INTEGER', 'PRIMARY KEY AUTOINCREMENT'],
    ['Genres', 'Name', 'TEXT', 'NOT NULL UNIQUE'],
    ['Reviews', 'Id', 'INTEGER', 'PRIMARY KEY AUTOINCREMENT'],
    ['Reviews', 'MovieId', 'INTEGER', 'NOT NULL UNIQUE, FOREIGN KEY REFERENCES Movies(Id) ON DELETE CASCADE'],
    ['Reviews', 'Rating', 'REAL', 'NOT NULL'],
    ['Reviews', 'Comment', 'TEXT', ''],
    ['Reviews', 'DateWatched', 'TEXT', 'NOT NULL'],
    ['Statistics', 'Id', 'INTEGER', 'PRIMARY KEY CHECK (Id = 1)'],
    ['Statistics', 'TotalMovies', 'INTEGER', 'NOT NULL'],
    ['Statistics', 'WatchedCount', 'INTEGER', 'NOT NULL'],
    ['Statistics', 'AverageRating', 'REAL', 'NOT NULL'],
    ['Statistics', 'LastUpdated', 'TEXT', 'NOT NULL'],
])
a = insert_par_after(a, 'Хранение отзывов выделено в отдельную таблицу с уникальным внешним ключом MovieId, что '
                        'гарантирует не более одного отзыва на фильм и обеспечивает каскадное удаление отзыва вместе '
                        'с фильмом. Таблица Statistics содержит ровно одну запись (ограничение CHECK (Id = 1)), '
                        'которая перезаписывается при каждом изменении данных. Внешние ключи включаются прагмой '
                        'PRAGMA foreign_keys = ON при каждом открытии соединения.')

# ----------------------------------------------------------------------------
# РАЗДЕЛ 5. ДИАГРАММЫ КЛАССОВ
# ----------------------------------------------------------------------------
a = el_h5
a = insert_par_after(a, 'Программная система состоит из трёх сборкам: Model, ViewModel и View. Диаграммы классов '
                        'изображены на рисунках 5.1–5.3.')
a = insert_caption_after(a, 'Рисунок 5.1 – Диаграмма классов сущностей предметной области (Movie, Director, Genre, '
                             'Review, AppStatistics, IRepository, ConcreteRepository)')
a = insert_caption_after(a, 'Рисунок 5.2 – Диаграмма классов слоя представления моделей (MainWindowViewModel, '
                             'MoviesDataStore, ViewModelNavigator, INavigableViewModel, RelayCommand и модели '
                             'представления разделов)')
a = insert_caption_after(a, 'Рисунок 5.3 – Диаграмма классов слоя представления (MainWindow и пользовательские '
                             'контролы, привязанные к моделям представления)')
a = insert_par_after(a, 'Класс MainWindowViewModel создаёт экземпляр репозитория и хранилища данных, формирует пять '
                        'моделей представления разделов и регистрирует их в навигаторе. Каждая модель раздела '
                        'реализует интерфейс INavigableViewModel и получает общие данные через MoviesDataStore, '
                        'что исключает дублирование источников данных между разделами.')

# ----------------------------------------------------------------------------
# РАЗДЕЛ 6. ОПИСАНИЕ КЛАССОВ ПРИЛОЖЕНИЯ
# ----------------------------------------------------------------------------
classes_rows = [
    ['Класс', 'Описание', 'Причина создания класса'],
    ['Movie', 'Основная модель фильма', 'Хранение центральной сущности предметной области'],
    ['Director', 'Модель режиссёра', 'Хранение ФИО режиссёра и исключение дублей в справочнике'],
    ['Genre', 'Модель жанра', 'Хранение названия жанра и исключение дублей в справочнике'],
    ['Review', 'Модель отзыва', 'Хранение оценки, комментария и даты просмотра фильма'],
    ['AppStatistics', 'Модель статистики', 'Хранение агрегированных показателей коллекции'],
    ['IRepository', 'Интерфейс репозитория', 'Абстракция доступа к данным и подмена данных в дизайнере WPF'],
    ['ConcreteRepository', 'Класс доступа к базе данных SQLite', 'Чтение, сохранение и удаление данных о фильмах'],
    ['MoviesDataStore', 'Централизованное хранилище данных', 'Общий источник данных для всех разделов приложения'],
    ['RelayCommand', 'Класс команды пользовательского интерфейса', 'Связь кнопок интерфейса с методами ViewModel'],
    ['INavigableViewModel', 'Интерфейс навигируемого представления', 'Единый способ управления видимостью разделов'],
    ['ViewModelNavigator', 'Класс навигации между разделами', 'Переключение активного раздела по одному вызову'],
    ['MainWindowViewModel', 'Главная модель представления окна', 'Управление переходами между разделами приложения'],
    ['MainMenuViewModel', 'Модель представления главного меню', 'Отображение стартового экрана приложения'],
    ['MoviesViewModel', 'Модель представления списка фильмов', 'Выбор, отображение и удаление записей коллекции'],
    ['AddMovieViewModel', 'Модель представления формы добавления', 'Создание нового фильма в базе данных'],
    ['MovieDetailsViewModel', 'Модель представления карточки фильма', 'Просмотр деталей и сохранение отзыва'],
    ['StatisticsViewModel', 'Модель представления статистики', 'Отображение показателей и управление тестовыми данными'],
]
a = el_h6
a = insert_par_after(a, 'Для понимания предназначения каждого объекта диаграммы классов программы необходимо описать '
                        'основные классы приложения. Их перечень приведён в таблице 6.1.')
a = insert_caption_after(a, 'Таблица 6.1 – Описание классов')
a = insert_table_after(a, classes_rows)

detail_tables = [
    ('Таблица 6.2 – Описание класса Movie', [
        ['Элемент', 'Тип', 'Описание'],
        ['Id', 'Свойство', 'Идентификатор фильма'],
        ['Title', 'Свойство', 'Название фильма'],
        ['ReleaseYear', 'Свойство', 'Год выхода фильма'],
        ['Description', 'Свойство', 'Описание фильма'],
        ['IsTestData', 'Свойство', 'Признак тестовой записи'],
        ['DirectorId', 'Свойство', 'Идентификатор режиссёра'],
        ['Director', 'Свойство', 'Ссылка на объект режиссёра'],
        ['GenreId', 'Свойство', 'Идентификатор жанра'],
        ['Genre', 'Свойство', 'Ссылка на объект жанра'],
        ['Review', 'Свойство', 'Ссылка на отзыв о фильме'],
    ]),
    ('Таблица 6.3 – Описание класса ConcreteRepository', [
        ['Элемент', 'Тип', 'Описание'],
        ['_connectionString', 'Поле', 'Строка подключения к локальной базе данных SQLite'],
        ['ConcreteRepository(string)', 'Конструктор', 'Запоминает путь к БД и вызывает её инициализацию'],
        ['OpenConnection()', 'Метод', 'Открывает соединение и включает внешние ключи'],
        ['InitializeDatabase()', 'Метод', 'Создаёт таблицы и заполняет базу тестовыми данными'],
        ['DropLegacyMovieTable(...)', 'Метод', 'Удаляет таблицу Films устаревшей структуры'],
        ['GetAllMovies()', 'Метод', 'Возвращает фильмы вместе с режиссёрами, жанрами и отзывами'],
        ['GetAllGenres()', 'Метод', 'Возвращает список жанров'],
        ['GetAllDirectors()', 'Метод', 'Возвращает список режиссёров'],
        ['GetAllReviews()', 'Метод', 'Возвращает список отзывов'],
        ['AddMovie(Movie)', 'Метод', 'Добавляет фильм и связанный отзыв в одной транзакции'],
        ['UpdateMovie(Movie)', 'Метод', 'Обновляет существующую запись фильма'],
        ['DeleteMovie(Movie)', 'Метод', 'Удаляет фильм (отзыв удаляется каскадно)'],
        ['SaveReview(Review)', 'Метод', 'Добавляет или обновляет отзыв о фильме'],
        ['EnsureDirector / EnsureGenre', 'Метод', 'Находит или создаёт запись в справочниках'],
        ['ClearGeneratedData()', 'Метод', 'Удаляет только тестовые фильмы'],
        ['ClearAll()', 'Метод', 'Полностью очищает базу данных'],
        ['GetStatistics() / SaveStatistics()', 'Метод', 'Читает и сохраняет агрегированную статистику'],
    ]),
    ('Таблица 6.4 – Описание класса MoviesDataStore', [
        ['Элемент', 'Тип', 'Описание'],
        ['_repository', 'Поле', 'Репозиторий для доступа к базе данных'],
        ['_selectedMovie', 'Поле', 'Текущий выбранный фильм'],
        ['_statistics', 'Поле', 'Текущие значения статистики'],
        ['Movies', 'Свойство', 'Коллекция фильмов для связанных списков'],
        ['Genres', 'Свойство', 'Коллекция жанров'],
        ['Directors', 'Свойство', 'Коллекция режиссёров'],
        ['Reviews', 'Свойство', 'Коллекция отзывов'],
        ['SelectedMovie', 'Свойство', 'Выбранный фильм с уведомлением подписчиков'],
        ['Statistics', 'Свойство', 'Показатели коллекции'],
        ['Refresh()', 'Метод', 'Перезагружает данные из репозитория и сохраняет статистику'],
        ['AddMovie(Movie)', 'Метод', 'Добавляет фильм и обновляет хранилище'],
        ['DeleteSelectedMovie()', 'Метод', 'Удаляет выбранный фильм'],
        ['SaveReview(Review)', 'Метод', 'Сохраняет отзыв о выбранном фильме'],
        ['GenerateTestData()', 'Метод', 'Заполняет базу тестовыми фильмами'],
        ['ClearGeneratedData()', 'Метод', 'Удаляет сгенерированные тестовые данные'],
        ['ClearAll()', 'Метод', 'Очищает все таблицы базы данных'],
        ['SaveStatisticsToDb()', 'Метод', 'Пересчитывает и сохраняет статистику'],
    ]),
    ('Таблица 6.5 – Описание класса RelayCommand', [
        ['Элемент', 'Тип', 'Описание'],
        ['_execute', 'Поле', 'Действие, выполняемое при вызове команды'],
        ['_canExecute', 'Поле', 'Условие, определяющее возможность выполнения команды'],
        ['RelayCommand(Action, Func?)', 'Конструктор', 'Создаёт объект команды без параметра'],
        ['RelayCommand(Action<object>, ...)', 'Конструктор', 'Создаёт объект команды с параметром'],
        ['CanExecuteChanged', 'Событие', 'Событие изменения доступности команды'],
        ['CanExecute(object?)', 'Метод', 'Проверяет, можно ли выполнить команду'],
        ['Execute(object?)', 'Метод', 'Выполняет связанную команду'],
        ['RaiseCanExecuteChanged()', 'Метод', 'Принудительно обновляет состояние команд'],
    ]),
    ('Таблица 6.6 – Описание класса MainWindowViewModel', [
        ['Элемент', 'Тип', 'Описание'],
        ['_navigator', 'Поле', 'Навигатор между моделями представления разделов'],
        ['NullRepository', 'Вложенный класс', 'Пустая реализация IRepository для дизайнера WPF'],
        ['MainMenuContent', 'Свойство', 'Модель представления главного меню'],
        ['MoviesContent', 'Свойство', 'Модель представления списка фильмов'],
        ['AddMovieContent', 'Свойство', 'Модель представления формы добавления фильма'],
        ['MovieDetailsContent', 'Свойство', 'Модель представления деталей и отзыва'],
        ['StatisticsContent', 'Свойство', 'Модель представления статистики'],
        ['MainMenuBtnCommand', 'Свойство', 'Команда перехода к главному меню'],
        ['MoviesBtnCommand', 'Свойство', 'Команда перехода к списку фильмов'],
        ['AddMovieBtnCommand', 'Свойство', 'Команда перехода к форме добавления'],
        ['MovieDetailsBtnCommand', 'Свойство', 'Команда перехода к карточке фильма'],
        ['StatisticsBtnCommand', 'Свойство', 'Команда перехода к статистике'],
        ['MainWindowViewModel(IRepository)', 'Конструктор', 'Создаёт хранилище, представления и команды'],
        ['GoToMainMenu() и др.', 'Метод', 'Переключают активный раздел через навигатор'],
        ['OnPropertyChanged(...)', 'Метод', 'Уведомляет интерфейс об изменении свойств'],
    ]),
]
for cap, rows in detail_tables:
    a = insert_caption_after(a, cap)
    a = insert_table_after(a, rows)

# ----------------------------------------------------------------------------
# РАЗДЕЛ 7. РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ
# ----------------------------------------------------------------------------
a = el_h7
a = insert_par_after(a, 'После запуска приложения пользователь открывает главное окно программы «Кинодневник». '
                        'Проект состоит из четырёх основных частей: Model (классы предметной области и репозиторий), '
                        'ViewModel (модели представления и навигация), View (окно и пользовательские контролы) и '
                        'локальной базы данных SQLite. Окно разделено на две области: слева расположено навигационное '
                        'меню с пятью кнопками, справа — область содержимого активного раздела.')
a = insert_par_after(a, 'При первом запуске программа автоматически создаёт файл базы данных moviediary.db и заполняет '
                        'его четырьмя тестовыми фильмами («Побег из Шоушенка», «Начало», «Интерстеллар», «Матрица»). '
                        'В разделе «Главное меню» отображается приветственный экран с предложением выбрать раздел.')
a = insert_par_after(a, 'В разделе «Фильмы» отображается список всех записей коллекции, загруженных из базы данных. '
                        'Пользователь выделяет нужную строку списка; выбранный фильм запоминается общим хранилищем '
                        'данных и доступен в других разделах. Кнопка «Удалить выбранный фильм» активна только тогда, '
                        'когда в списке выбран хотя бы один фильм.')
a = insert_par_after(a, 'В разделе «Добавить фильм» расположена форма с полями «Название фильма», «Год», «Жанр», '
                        '"Режиссер" и «Описание». Кнопка «Добавить фильм» становится доступной после ввода названия. '
                        'После сохранения форма очищается, а новый фильм появляется в списке раздела «Фильмы». Если '
                        'жанр или режиссёр ранее не встречались, программа автоматически создаёт записи в справочниках.')
a = insert_par_after(a, 'Для просмотра полной информации о фильме пользователь сначала выделяет запись в списке, '
                        'а затем переходит в раздел «Детали и отзыв». В верхней части раздела отображаются название, '
                        'год, жанр, режиссёр и описание выбранного фильма. Ниже расположены поля «Оценка (1-10)» и '
                        '«Ваш отзыв»: после нажатия кнопки «Сохранить отзыв» данные записываются в базу, а фильм '
                        'переходит в разряд просмотренных.')
a = insert_par_after(a, 'В разделе «Статистика и данные» выводятся количество фильмов, количество просмотренных '
                        'фильмов, средняя оценка, а также число жанров, режиссёров и отзывов. Здесь же доступны кнопки '
                        '«Сгенерировать данные», «Удалить тестовые» и «Удалить ВСЕ», управляющие содержимым базы данных.')
a = insert_caption_after(a, 'Рисунок 7.1 – Главное окно приложения «Кинодневник» (раздел «Главное меню»)')
a = insert_caption_after(a, 'Рисунок 7.2 – Раздел «Фильмы» со списком записей коллекции')
a = insert_caption_after(a, 'Рисунок 7.3 – Форма добавления фильма в раздел «Добавить фильм»')
a = insert_caption_after(a, 'Рисунок 7.4 – Карточка фильма с оценкой и отзывом в разделе «Детали и отзыв»')
a = insert_caption_after(a, 'Рисунок 7.5 – Раздел «Статистика и данные»')

# ----------------------------------------------------------------------------
# РАЗДЕЛ 8. КОД ПРИЛОЖЕНИЯ
# ----------------------------------------------------------------------------
code_sections = [
    ('Код класса – модели фильма:', '''public class Movie
{
    public int Id { get; set; }
    public string Title { get; set; } = string.Empty;
    public int ReleaseYear { get; set; }
    public string Description { get; set; } = string.Empty;
    public bool IsTestData { get; set; }

    public int? DirectorId { get; set; }
    public Director? Director { get; set; }

    public int? GenreId { get; set; }
    public Genre? Genre { get; set; }

    public Review? Review { get; set; }
}'''),
    ('Код классов – моделей режиссёра, жанра и отзыва:', '''public class Director
{
    public int Id { get; set; }
    public string FullName { get; set; } = string.Empty;

    public override string ToString() => FullName;
}

public class Genre
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;

    public override string ToString() => Name;
}

public class Review
{
    public int Id { get; set; }
    public int MovieId { get; set; }
    public Movie? Movie { get; set; }
    public double Rating { get; set; }
    public string Comment { get; set; } = string.Empty;
    public DateTime DateWatched { get; set; }
}'''),
    ('Код класса – модели статистики:', '''public class AppStatistics
{
    public int TotalMovies { get; set; }
    public int WatchedCount { get; set; }
    public double AverageRating { get; set; }
    public DateTime LastUpdated { get; set; }
}'''),
    ('Код интерфейса – репозитория:', '''public interface IRepository
{
    List<Movie> GetAllMovies();
    List<Genre> GetAllGenres();
    List<Director> GetAllDirectors();
    List<Review> GetAllReviews();

    void AddMovie(Movie movie);
    void UpdateMovie(Movie movie);
    void DeleteMovie(Movie movie);
    void SaveReview(Review review);

    void ClearAll();
    void ClearGeneratedData();
    void GenerateTestData();

    AppStatistics GetStatistics();
    void SaveStatistics(AppStatistics stats);
}'''),
    ('Код класса – репозитория для работы с базой данных (фрагмент):', '''public class ConcreteRepository : IRepository
{
    private readonly string _connectionString;

    public ConcreteRepository(string dbPath = "moviediary.db")
    {
        _connectionString = $"Data Source={dbPath}";
        InitializeDatabase();
    }

    private SqliteConnection OpenConnection()
    {
        var connection = new SqliteConnection(_connectionString);
        connection.Open();
        using var pragma = new SqliteCommand("PRAGMA foreign_keys = ON;", connection);
        pragma.ExecuteNonQuery();
        return connection;
    }

    private void InitializeDatabase()
    {
        using var connection = OpenConnection();
        DropLegacyMovieTable(connection);

        // CREATE TABLE IF NOT EXISTS Directors, Genres, Movies, Reviews, Statistics
        // ...

        using var countCommand = new SqliteCommand("SELECT COUNT(*) FROM Movies;", connection);
        var movieCount = Convert.ToInt32(countCommand.ExecuteScalar());
        if (movieCount == 0)
        {
            connection.Close();
            GenerateTestData();
        }
    }

    public void AddMovie(Movie movie)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();

        movie.DirectorId = EnsureDirector(connection, transaction, movie.Director);
        movie.GenreId = EnsureGenre(connection, transaction, movie.Genre);

        var query = @"
            INSERT INTO Movies (Title, ReleaseYear, Description, DirectorId, GenreId, IsTestData)
            VALUES (@Title, @ReleaseYear, @Description, @DirectorId, @GenreId, @IsTestData);
            SELECT last_insert_rowid();";

        using (var command = new SqliteCommand(query, connection, transaction))
        {
            command.Parameters.AddWithValue("@Title", movie.Title);
            command.Parameters.AddWithValue("@ReleaseYear", movie.ReleaseYear);
            command.Parameters.AddWithValue("@Description", (object?)movie.Description ?? DBNull.Value);
            command.Parameters.AddWithValue("@DirectorId", (object?)movie.DirectorId ?? DBNull.Value);
            command.Parameters.AddWithValue("@GenreId", (object?)movie.GenreId ?? DBNull.Value);
            command.Parameters.AddWithValue("@IsTestData", movie.IsTestData ? 1 : 0);
            movie.Id = Convert.ToInt32(command.ExecuteScalar());
        }

        if (movie.Review != null)
        {
            movie.Review.MovieId = movie.Id;
            SaveReview(connection, transaction, movie.Review);
        }

        transaction.Commit();
    }
}'''),
    ('Код класса – команды:', '''public class RelayCommand : ICommand
{
    private readonly Action<object> _execute;
    private readonly Func<object, bool>? _canExecute;

    public RelayCommand(Action execute, Func<bool>? canExecute = null)
        : this(
            ignoredParameter => execute(),
            canExecute == null ? null : ignoredParameter => canExecute())
    {
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter)
    {
        return _canExecute == null || _canExecute(parameter!);
    }

    public void Execute(object? parameter)
    {
        _execute(parameter!);
    }

    public void RaiseCanExecuteChanged()
    {
        CommandManager.InvalidateRequerySuggested();
    }
}'''),
    ('Код интерфейса и класса навигации:', '''public interface INavigableViewModel
{
    bool IsVisible { get; set; }
}

public class ViewModelNavigator
{
    private readonly IReadOnlyList<INavigableViewModel> _pages;

    public ViewModelNavigator(IReadOnlyList<INavigableViewModel> pages)
    {
        _pages = pages ?? throw new ArgumentNullException(nameof(pages));
    }

    public void NavigateTo(INavigableViewModel target)
    {
        if (target == null)
            throw new ArgumentNullException(nameof(target));

        foreach (var page in _pages)
            page.IsVisible = ReferenceEquals(page, target);
    }
}'''),
    ('Код класса – главной модели представления окна (фрагмент):', '''public class MainWindowViewModel : INotifyPropertyChanged
{
    private readonly ViewModelNavigator _navigator;

    public MainWindowViewModel()
        : this(DesignerProperties.GetIsInDesignMode(new DependencyObject())
            ? new NullRepository()
            : new ConcreteRepository())
    {
    }

    public MainWindowViewModel(IRepository repository)
    {
        ArgumentNullException.ThrowIfNull(repository);

        var dataStore = new MoviesDataStore(repository);

        MainMenuContent = new MainMenuViewModel();
        MoviesContent = new MoviesViewModel(dataStore);
        AddMovieContent = new AddMovieViewModel(dataStore);
        MovieDetailsContent = new MovieDetailsViewModel(dataStore);
        StatisticsContent = new StatisticsViewModel(dataStore);

        _navigator = new ViewModelNavigator(new INavigableViewModel[]
        {
            MainMenuContent,
            MoviesContent,
            AddMovieContent,
            MovieDetailsContent,
            StatisticsContent
        });

        MainMenuBtnCommand = new RelayCommand(GoToMainMenu);
        MoviesBtnCommand = new RelayCommand(GoToMovies);
        AddMovieBtnCommand = new RelayCommand(GoToAddMovie);
        MovieDetailsBtnCommand = new RelayCommand(GoToMovieDetails);
        StatisticsBtnCommand = new RelayCommand(GoToStatistics);

        GoToMainMenu();
    }

    private void GoToMovies()
    {
        _navigator.NavigateTo(MoviesContent);
    }
}'''),
    ('Код класса – centralized хранилища данных MoviesDataStore (фрагмент):', '''public class MoviesDataStore : INotifyPropertyChanged
{
    private readonly IRepository _repository;
    private Movie? _selectedMovie;
    private AppStatistics _statistics = new() { LastUpdated = DateTime.Now };

    public MoviesDataStore(IRepository repository)
    {
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        Movies = new ObservableCollection<Movie>();
        Genres = new ObservableCollection<Genre>();
        Directors = new ObservableCollection<Director>();
        Reviews = new ObservableCollection<Review>();
        Refresh();
    }

    public ObservableCollection<Movie> Movies { get; }
    public Movie? SelectedMovie
    {
        get => _selectedMovie;
        set
        {
            if (ReferenceEquals(_selectedMovie, value))
                return;

            _selectedMovie = value;
            OnPropertyChanged();
        }
    }

    public void Refresh()
    {
        var selectedId = SelectedMovie?.Id;

        Movies.Clear();
        foreach (var movie in _repository.GetAllMovies())
            Movies.Add(movie);

        SelectedMovie = selectedId.HasValue
            ? Movies.FirstOrDefault(m => m.Id == selectedId.Value)
            : null;

        SaveStatisticsToDb();
    }

    private void SaveStatisticsToDb()
    {
        var ratings = Movies
            .Where(m => m.Review != null && m.Review.Rating > 0)
            .Select(m => m.Review!.Rating)
            .ToList();

        var stats = new AppStatistics
        {
            TotalMovies = Movies.Count,
            WatchedCount = Movies.Count(m => m.Review != null),
            AverageRating = ratings.Count == 0 ? 0 : ratings.Average(),
            LastUpdated = DateTime.Now
        };

        _repository.SaveStatistics(stats);
        Statistics = _repository.GetStatistics();
    }
}'''),
    ('Код класса – модели представления добавления фильма (фрагмент):', '''public class AddMovieViewModel : INotifyPropertyChanged, INavigableViewModel
{
    private readonly MoviesDataStore _dataStore;

    public AddMovieViewModel(MoviesDataStore dataStore)
    {
        _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));
        AddNewMovieCommand = new RelayCommand(AddNewMovie, CanAddNewMovie);
    }

    public string NewTitle { get; set; } = string.Empty;
    public int NewYear { get; set; } = DateTime.Now.Year;
    public string NewGenre { get; set; } = string.Empty;
    public string NewDirector { get; set; } = string.Empty;
    public string NewDescription { get; set; } = string.Empty;

    public ICommand AddNewMovieCommand { get; }

    private void AddNewMovie()
    {
        var movie = new Movie
        {
            Title = NewTitle,
            ReleaseYear = NewYear,
            Description = NewDescription,
            Genre = string.IsNullOrWhiteSpace(NewGenre) ? null : new Genre { Name = NewGenre.Trim() },
            Director = string.IsNullOrWhiteSpace(NewDirector) ? null : new Director { FullName = NewDirector.Trim() }
        };

        _dataStore.AddMovie(movie);

        NewTitle = string.Empty;
        NewGenre = string.Empty;
        NewDirector = string.Empty;
        NewDescription = string.Empty;
        NewYear = DateTime.Now.Year;
    }

    private bool CanAddNewMovie()
    {
        return !string.IsNullOrWhiteSpace(NewTitle);
    }
}'''),
    ('Код класса – модели представления детальной информации (фрагмент):', '''public class MovieDetailsViewModel : INotifyPropertyChanged, INavigableViewModel
{
    private readonly MoviesDataStore _dataStore;

    public MovieDetailsViewModel(MoviesDataStore dataStore)
    {
        _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));
        _dataStore.PropertyChanged += OnDataStorePropertyChanged;

        SaveReviewCommand = new RelayCommand(SaveReview, CanSaveReview);
        SyncEditFieldsFromSelection();
    }

    public Movie? SelectedMovie => _dataStore.SelectedMovie;
    public double EditRating { get; set; }
    public string EditReview { get; set; } = string.Empty;

    public ICommand SaveReviewCommand { get; }

    private void SaveReview()
    {
        if (SelectedMovie == null)
            return;

        var review = SelectedMovie.Review ?? new Review { MovieId = SelectedMovie.Id };
        review.MovieId = SelectedMovie.Id;
        review.Rating = EditRating;
        review.Comment = EditReview ?? string.Empty;
        review.DateWatched = DateTime.Now;
        review.Movie = SelectedMovie;

        _dataStore.SaveReview(review);
    }

    private bool CanSaveReview()
    {
        return SelectedMovie != null;
    }
}'''),
    ('Код класса – модели представления статистики (фрагмент):', '''public class StatisticsViewModel : INotifyPropertyChanged, INavigableViewModel
{
    private readonly MoviesDataStore _dataStore;

    public StatisticsViewModel(MoviesDataStore dataStore)
    {
        _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));

        GenerateDataCommand = new RelayCommand(GenerateData);
        ClearGeneratedDataCommand = new RelayCommand(ClearGeneratedData);
        ClearAllDataCommand = new RelayCommand(ClearAllData);
    }

    public AppStatistics Statistics => _dataStore.Statistics;
    public ObservableCollection<Genre> Genres => _dataStore.Genres;
    public ObservableCollection<Director> Directors => _dataStore.Directors;
    public ObservableCollection<Review> Reviews => _dataStore.Reviews;

    public ICommand GenerateDataCommand { get; }
    public ICommand ClearGeneratedDataCommand { get; }
    public ICommand ClearAllDataCommand { get; }

    private void GenerateData()
    {
        _dataStore.GenerateTestData();
    }
}'''),
    ('Код фрагмента XAML главного окна (привязка разделов и кнопок навигации):', '''<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="180"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <StackPanel Grid.Column="0" Background="LightGray" Margin="5">
        <Button Command="{Binding MainMenuBtnCommand}" Content="Главное меню" Margin="5"/>
        <Button Command="{Binding MoviesBtnCommand}" Content="Фильмы" Margin="5"/>
        <Button Command="{Binding AddMovieBtnCommand}" Content="Добавить фильм" Margin="5"/>
        <Button Command="{Binding MovieDetailsBtnCommand}" Content="Детали и отзыв" Margin="5"/>
        <Button Command="{Binding StatisticsBtnCommand}" Content="Статистика" Margin="5"/>
    </StackPanel>

    <Border Grid.Column="1" Background="White" BorderBrush="Gray" BorderThickness="1" Margin="5">
        <Grid Margin="10">
            <local:MoviesUserControl
                DataContext="{Binding MoviesContent}"
                Visibility="{Binding IsVisible, Converter={StaticResource BoolToVisibility}}"/>
            <local:AddMovieUserControl
                DataContext="{Binding AddMovieContent}"
                Visibility="{Binding IsVisible, Converter={StaticResource BoolToVisibility}}"/>
        </Grid>
    </Border>
</Grid>'''),
]

a = el_h8
a = insert_par_after(a, 'Основные процессы приложения реализованы в сборках Model, ViewModel и View. Далее будут '
                        'приведены основные классы и методы, определяющие работу программы «Кинодневник».')
for caption, code in code_sections:
    a = insert_par_after(a, caption, style='List Paragraph')
    a = insert_code_after(a, code)

# ----------------------------------------------------------------------------
# РАЗДЕЛ 9. ТЕСТИРОВАНИЕ (перед заключением)
# ----------------------------------------------------------------------------
a = insert_heading_after(el_zakl, '9', 'Тестирование разработанного приложения')
a = insert_par_after(a, 'Тестирование приложения проводилось путем последовательной проверки основных пользовательских '
                        'сценариев. В первую очередь проверялась работоспособность приложения при первом запуске: '
                        'программа создавала файл базы данных moviediary.db, формировала таблицы Directors, Genres, '
                        'Movies, Reviews и Statistics и автоматически заполняла коллекцию четырьмя тестовыми фильмами.')
a = insert_par_after(a, 'Следующим этапом было протестировано отображение списка фильмов. После перехода в раздел '
                        '«Фильмы» список должен содержать все записи базы данных, причём названия отображаются '
                        'корректно даже при отсутствии заполненных описаний. Проверка выделения элемента списка '
                        'подтвердила, что выбранный фильм становится доступен в разделе «Детали и отзыв».')
a = insert_par_after(a, 'Отдельно проверялись операции добавления и удаления данных. Новый фильм создавался через форму '
                        '«Добавить фильм»; при пустом поле названия кнопка добавления оставалась недоступной, а после '
                        'ввода названия запись появлялась в списке. Для несуществующего в справочниках жанра и '
                        'режиссёра программа корректно создавала новые записи без дублирования уже имеющихся. '
                        'Удаление выбранного фильма приводило к каскадному удалению связанного отзыва.')
a = insert_par_after(a, 'Кроме того, выполнялась проверка сохранения отзыва и накопления статистики. После выставления '
                        'оценки и сохранения комментария показатель «Просмотрено» увеличивался, а средняя оценка '
                        'пересчитывалась. Повторное сохранение отзыва для того же фильма не создавало дубликат, '
                        'поскольку операция выполнена как upsert по уникальному полю MovieId.')
a = insert_par_after(a, 'Проверялось и поведение приложения при неполных данных: фильм без режиссёра, без жанра и без '
                        'отзыва отображается в интерфейсе корректно, а статистика формируется без ошибок деления на ноль.')
a = insert_caption_after(a, 'Рисунок 9.1 – Результат проверки работы приложения «Кинодневник» и данных базы SQLite')
a = insert_par_after(a, 'Итоговое тестирование показало, что приложение «Кинодневник» корректно выполняет основные '
                        'функции домашней фильмотеки: загружает и отображает данные, добавляет и удаляет фильмы, '
                        'сохраняет отзывы и статистику, а также управляет тестовыми данными. Все выявленные в ходе '
                        'проверки замечания были устранены.')

# ----------------------------------------------------------------------------
# ОБНОВЛЕНИЕ СОДЕРЖАНИЯ, ВВЕДЕНИЯ, ЗАКЛЮЧЕНИЯ И СПИСКА ЛИТЕРАТУРЫ
# ----------------------------------------------------------------------------
# Содержание: заменяем старый TOC на поле
from docx.text.paragraph import Paragraph
toc_host = Paragraph(el_toc, d)
# удалить sdt-блок старого TOC
for el in list(body):
    if el.tag == qn('w:sdt'):
        body.remove(el)
make_toc_field(toc_host)
toc_host.paragraph_format.page_break_before = True

# ВВЕДЕНИЕ — переписываем абзацы между заголовком «ВВЕДЕНИЕ» и разделом 1
el_vv = el_intro.getprevious()
while el_vv is not None and not (''.join(x.text for x in el_vv.iter(qn('w:t'))).strip().startswith('ВВЕДЕНИЕ')):
    el_vv = el_vv.getprevious()
new_intro = [
    'Курсовой проект – работа, представленная ниже, посвящена разработке десктопного Windows-приложения «Кинодневник» '
    'с использованием технологии WPF на языке программирования C# платформы .NET 8. Целью проекта является разработка '
    'и реализация функциональности, которая включает в себя сохранение статистических параметров в базе данных, '
    'возможность удобного редактирования БД (включая CRUD-операции), автоматическое заполнение БД тестовыми данными '
    'и предоставление удобного пользовательского интерфейса.',
    'Данный курсовой проект предполагает индивидуальное выполнение, что способствует развитию компетенций в области '
    'проектирования информационных технологий. Он помогает оценить надежность и качество функционирования объекта '
    'проектирования, разработать базовые и прикладные информационные технологии, а также научиться реализовывать '
    'средства информационных технологий, используя методические, информационные, математические, алгоритмические, '
    'технические и программные подходы.',
    'Важным аспектом проекта является готовность к сборке информационной системы из готовых компонентов. Решение '
    'разделено на проекты Model, ViewModel и View, что позволяет рассматривать каждый слой как самостоятельный '
    'компонент и собирать систему из независимых частей. Кроме того, проект ставит своей целью развить навыки '
    'составления инструкций по эксплуатации информационных систем, что является важным аспектом в процессе внедрения '
    'и использования разработанных систем.',
    'Архитектурной основой приложения послужил шаблон проектирования MVVM (Model – View – ViewModel), который '
    'обеспечил отделение пользовательского интерфейса от бизнес-логики и работы с базой данных. Взаимодействие с '
    'интерфейсом организовано командным подходом через собственный класс RelayCommand, а хранение данных реализовано '
    'в реляционной СУБД SQLite через класс-репозиторий ConcreteRepository.',
    'В итоге, выполнение данного курсового проекта поможет закрепить и применить полученные знания и навыки, а также '
    'развить компетенции в области проектирования и разработки информационных технологий.',
    'В данном курсовом проекте используется программная среда Visual Studio 2022 Community языка программирования — C#.',
]
anchor = el_vv
for t in new_intro:
    anchor = insert_par_after(anchor, t)
# удалить старые абзацы введения (между новым текстом и разделом 1)
remove_between(anchor, el_h1)

# ЗАКЛЮЧЕНИЕ — переписываем текст между заголовком и списком литературы
new_concl = [
    'В ходе выполнения курсового проекта было разработано десктопное Windows-приложение «Кинодневник» с использованием '
    'технологии WPF и языка программирования C#. Приложение предназначено для ведения домашней фильмотеки и обеспечивает '
    'хранение сведений о фильмах, жанрах и режиссёрах, а также оценок и отзывов пользователя.',
    'В процессе выполнения работы была спроектирована и реализована база данных SQLite, содержащая таблицы фильмов, '
    'режиссёров, жанров, отзывов и статистики. Были настроены связи «один ко многим» и «один к одному», включены внешние '
    'ключи и каскадное удаление, а также реализовано автоматическое создание и первоначальное заполнение базы данных '
    'тестовыми фильмами при первом запуске.',
    'При разработке программы использовался архитектурный подход MVVM, который обеспечил разделение пользовательского '
    'интерфейса, логики представления и доступа к данным. Централизованное хранилище MoviesDataStore позволило нескольким '
    'разделам работать с одними и теми же данными, а навигатор ViewModelNavigator и флаг IsVisible обеспечили переключение '
    'пяти разделов приложения: главного меню, списка фильмов, формы добавления, карточки фильма и статистики.',
    'Выполнение курсового проекта позволило закрепить практические навыки проектирования программных систем, работы с '
    'WPF-приложениями, средствами XAML-разметки и привязки данных, а также навыки работы с реляционными базами данных и '
    'паттерном «репозиторий». Были получены практические результаты по тестированию основных пользовательских сценариев.',
    'Таким образом, поставленная цель курсового проекта была достигнута. Разработанное приложение «Кинодневник» '
    'представляет собой функциональное средство ведения домашней фильмотеки, соответствует выданному заданию и может '
    'быть расширено новыми разделами и полями без изменения существующей архитектуры.',
]
first_after = d.paragraphs[zakl_idx + 1]._p
anchor = d.paragraphs[zakl_idx]._p
for t in new_concl:
    anchor = insert_par_after(anchor, t)
# удалить старые абзацы заключения (до СПИСКА ЛИТЕРАТУРЫ)
lit_idx = None
for i, p in enumerate(d.paragraphs):
    if 'СПИСОК ЛИТЕРАТУРЫ' in p.text and p.style.name == 'Мой заголовок':
        lit_idx = i
        break
stop = d.paragraphs[lit_idx]._p
cur = first_after
while cur is not None and cur is not stop:
    nxt = cur.getnext()
    if cur.tag == qn('w:p') and not ''.join(x.text for x in cur.iter(qn('w:t'))).strip():
        cur = nxt
        continue
    if cur.tag == qn('w:p'):
        cur.getparent().remove(cur)
        cur = nxt
    else:
        cur = nxt

# СПИСОК ЛИТЕРатуры — дополняем ссылками на документацию .NET/WPF
lit_idx = None
for i, p in enumerate(d.paragraphs):
    if 'СПИСОК ЛИТЕРАТУРЫ' in p.text and p.style.name == 'Мой заголовок':
        lit_idx = i
        break
last = d.paragraphs[-1]
extra_sources = [
    'Документация Microsoft Learn. Pattern MVVM (Model-View-ViewModel). URL: https://learn.microsoft.com/dotnet/architecture/modern-web-app-architectures/mvvm-pattern-windows-store-apps',
    'Документация Microsoft Learn. Класс ICommand и интерфейс INotifyPropertyChanged. URL: https://learn.microsoft.com/dotnet/api/',
    'SQLite.org. Документация по языку запросов SQLite. URL: https://www.sqlite.org/lang.html',
]
for s in extra_sources:
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = 425
    p.add_run(s)

# ----------------------------------------------------------------------------
# Печать страниц в нижних колонтитулах (как в исходном документе) и сохранение
# ----------------------------------------------------------------------------
d.save(DST)
print('saved:', DST)
EOF_MARKER_NOT_USED = None
