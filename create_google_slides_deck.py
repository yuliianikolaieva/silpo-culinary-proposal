from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


ROOT = Path(__file__).parent
OUTPUT = ROOT / "Silpo_x_Bolt_Food_commercial_proposal.pptx"

GREEN = "0B8F50"
DARK = "14221A"
MUTED = "5F6C65"
LIGHT = "EDF8F1"
LINE = "D8E2DA"
AMBER = "FFF7E5"
WHITE = "FFFFFF"


def rgb(value):
    return RGBColor.from_string(value)


def fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)


def border(shape, color=LINE, width=1):
    shape.line.color.rgb = rgb(color)
    shape.line.width = Pt(width)


def background(slide, color=WHITE):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def text_box(slide, x, y, w, h, text="", size=12, color=DARK, bold=False,
             font="Arial", align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
             margin=0.08, italic=False):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    return shape


def card(slide, x, y, w, h, title, value=None, body=None, bg=WHITE, accent=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    fill(shape, bg)
    border(shape, "BEDEC9" if accent else LINE)
    shape.adjustments[0] = 0.08
    text_box(slide, x + 0.14, y + 0.14, w - 0.28, 0.25, title.upper(), 7.5, MUTED, True)
    if value:
        text_box(slide, x + 0.14, y + 0.45, w - 0.28, 0.42, value, 18, GREEN if accent else DARK, True)
    if body:
        text_box(slide, x + 0.14, y + (0.95 if value else 0.48), w - 0.28, h - (1.05 if value else 0.58), body, 8.5, MUTED)
    return shape


def section_title(slide, eyebrow, title):
    text_box(slide, 0.6, 0.38, 12.0, 0.22, eyebrow.upper(), 8, GREEN, True)
    text_box(slide, 0.6, 0.63, 12.0, 0.55, title, 23, DARK, True)


def footer(slide, page):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(7.15), Inches(12.13), Inches(0.01))
    fill(line, LINE)
    line.line.fill.background()
    text_box(slide, 0.6, 7.2, 8.8, 0.18, "Сільпо × Bolt Food · Commercial proposal · October 2026", 6.5, MUTED)
    text_box(slide, 12.0, 7.2, 0.7, 0.18, str(page), 6.5, MUTED, align=PP_ALIGN.RIGHT)


def add_chart(slide):
    data = CategoryChartData()
    data.categories = ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру"]
    data.add_series("Кулінарія", (0.4, 11.2, 20.2, 24.0, 29.5, 33.0, 38.8, 42.6, 43.1, 46.5, 45.0, 48.4))
    data.add_series("Алко + напої", (0, 0, 0, 0, 0, 0, 0, 0, 5.9, 6.3, 6.1, 6.6))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED, Inches(0.6), Inches(2.15), Inches(7.25), Inches(3.3), data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.maximum_scale = 60
    chart.value_axis.minimum_scale = 0
    chart.value_axis.tick_labels.font.size = Pt(7)
    chart.category_axis.tick_labels.font.size = Pt(7)
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = rgb(GREEN)
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.rgb = rgb("C9E74A")
    chart.plots[0].gap_width = 45
    return chart


def add_image(slide, image_name, x, y, w, h):
    path = ROOT / "assets" / image_name
    slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1
    slide = prs.slides.add_slide(blank)
    background(slide)
    section_title(slide, "Сценарій 1 · план товарообігу та вимоги до категорії",
                  "100 локацій · ready meals як окрема delivery-категорія")
    text_box(slide, 0.6, 1.25, 12.0, 0.35,
             "Базовий forecast побудований на даних Сільпо та фактичних продажах кулінарії retail-партнерів Bolt Food. Planning FX: €1 = ₴50.",
             9, MUTED)
    card(slide, 8.1, 1.85, 1.35, 1.1, "Базовий GMV", "€407.5k", "₴20.4m", LIGHT, True)
    card(slide, 9.65, 1.85, 1.35, 1.1, "Кулінарія", "€382.6k", "₴19.1m", WHITE)
    card(slide, 11.2, 1.85, 1.35, 1.1, "Замовлення", "30,610", "за 12 міс.", WHITE)
    add_chart(slide)
    text_box(slide, 0.65, 5.55, 7.1, 0.25,
             "Алкоголь і soft drinks закладені з вересня; за готовності можуть стартувати одразу або після 3-місячного пілота — GMV буде перерахований.",
             7.7, MUTED)
    card(slide, 8.1, 3.2, 4.45, 0.7, "Консервативний сценарій", "€349.7k / ₴17.5m", "Зрілий OSPD 1.00 · 26,379 orders", AMBER)
    card(slide, 8.1, 4.05, 4.45, 0.7, "Базовий сценарій", "€407.5k / ₴20.4m", "Зрілий OSPD 1.25 · 30,610 orders", LIGHT, True)
    card(slide, 8.1, 4.9, 4.45, 0.7, "Оптимістичний сценарій", "€581.1k / ₴29.1m", "Зрілий OSPD 2.00 · 43,303 orders", WHITE)
    text_box(slide, 8.1, 5.9, 4.3, 0.25, "ЩО ПОТРІБНО ДЛЯ КАТЕГОРІЇ", 8, GREEN, True)
    card(slide, 8.1, 6.18, 1.4, 0.75, "Кулінарія", None, "80 → 120 → 150 core SKU\nТоп-SKU в наявності", LIGHT, True)
    card(slide, 9.62, 6.18, 1.4, 0.75, "Алкоголь", None, "25 SKU\nlegal + age-check", WHITE)
    card(slide, 11.14, 6.18, 1.4, 0.75, "Soft drinks", None, "20 SKU\n30% attach target", WHITE)
    footer(slide, 1)

    # Slide 2
    slide = prs.slides.add_slide(blank)
    background(slide)
    section_title(slide, "Сценарій 2 · комерційна пропозиція",
                  "12% на stores + умовне зниження комісії Dark Kitchen")
    quote = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.35), Inches(12.1), Inches(0.72))
    fill(quote, LIGHT)
    border(quote, "BEDEC9")
    text_box(slide, 0.82, 1.52, 11.65, 0.35,
             "12% на stores — мінімальна робоча ставка: вона покриває courier cost та операційне виконання, поки Bolt інвестує €55.2k у запуск і зберігає co-funding DK/resto.",
             10, DARK, True)
    steps = [
        ("М1–3 · LAUNCH", "Stores: 12%", "DK працює за поточною ставкою 25%. Bolt керує запуском попиту та courier allocation."),
        ("КОНТРОЛЬНА ТОЧКА", "2,540 orders + €33k GMV", "Поріг stores за перші 3 місяці із погодженого базового rollout-плану."),
        ("М4–12 · ЯКЩО ДОСЯГНУТО", "DK: 25% → 20%", "Знижена ставка діє до кінця 12-місячного періоду."),
        ("ЯКЩО НЕ ДОСЯГНУТО", "DK: 25%", "Зниження не активується. Окремо погоджуємо recovery plan."),
    ]
    for i, (label, value, body) in enumerate(steps):
        x = 0.6 + i * 3.03
        card(slide, x, 2.35, 2.84, 1.6, label, value, body, LIGHT if i == 2 else WHITE, i == 2)
    card(slide, 0.6, 4.35, 3.8, 1.15, "Економія Сільпо · DK", "€34.0k / ₴1.70m",
         "М4–12: 5 в.п. (25% → 20%) для €679k / ₴34.0m DK товарообігу. Це 5% від GMV, на який діє зниження.", LIGHT, True)
    card(slide, 4.62, 4.35, 3.8, 1.15, "Еквівалент за повний рік", "€45.3k / ₴2.26m",
         "Показово: 5% × €905.3k DK annual GMV. Не є умовою цього offer.", WHITE)
    card(slide, 8.64, 4.35, 4.06, 1.15, "Комерційна логіка", "Спочатку результат",
         "Зниження DK — винагорода за доведений попит stores, а не upfront subsidy. Це захищає unit economics обох сторін.", AMBER)
    recommendation = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(5.9), Inches(12.1), Inches(0.74))
    fill(recommendation, AMBER)
    border(recommendation, "ECD9AB")
    text_box(slide, 0.82, 6.08, 11.65, 0.32,
             "Рекомендація для захисту прибутковості: для кулінарії stores допускаємо markup до 10% до ціни полиці Сільпо; alcohol та soft drinks погоджуємо окремо за category economics.",
             9, DARK)
    footer(slide, 2)

    # Slide 3
    slide = prs.slides.add_slide(blank)
    background(slide)
    section_title(slide, "Цінність для Сільпо",
                  "Інвестиції Bolt, savings Сільпо та маркетинг launch")
    card(slide, 0.6, 1.35, 2.8, 1.15, "Bolt Plus · 6 місяців", "€20.0k / ₴1.0m",
         "Stores €5.4k + DK/resto €14.6k. Plus-клієнти мають вищий чек і частіші замовлення; ціль +10–15% GMV uplift.", LIGHT, True)
    card(slide, 3.62, 1.35, 2.8, 1.15, "Free delivery · stores", "€14.0k",
         "6 із 12 місяців. Знімає delivery barrier у launch і стимулює перше та повторне замовлення.", WHITE)
    card(slide, 6.64, 1.35, 2.8, 1.15, "Public launch + inventory", "€21.2k",
         "€10.0k cash + €11.2k in-kind: banner, push/email, hero та modal.", WHITE)
    card(slide, 9.66, 1.35, 3.05, 1.15, "Нові інвестиції Bolt", "€55.2k / ₴2.76m",
         "13.5% від stores GMV; 4.2% від combined stores + DK GMV.", LIGHT, True)
    text_box(slide, 0.6, 2.85, 12.0, 0.24, "МАРКЕТИНГОВИЙ ПЛАН", 8, GREEN, True)
    marketing = [
        ("Bolt Plus", "Free delivery для лояльної аудиторії з вищим чеком."),
        ("Free delivery", "Знімає бар’єр першого замовлення у ready meals."),
        ("Sponsored Listing · DK", "30% discount на чинний spend для додаткової видимості."),
        ("Co-funding · DK/resto", "Чинна підтримка попиту €99.1k annual run-rate зберігається без змін."),
    ]
    for i, (title, body) in enumerate(marketing):
        x = 0.6 + i * 3.03
        card(slide, x, 3.15, 2.84, 0.8, title, None, body, LIGHT if i in (0, 2) else WHITE, i in (0, 2))
    text_box(slide, 0.6, 4.28, 12.0, 0.24, "ПРИКЛАДИ ВИДИМОСТІ В BOLT FOOD", 8, GREEN, True)
    add_image(slide, "launch-modal.png", 0.6, 4.58, 2.3, 1.85)
    add_image(slide, "home-hero.png", 3.15, 4.58, 2.3, 1.85)
    add_image(slide, "in-app-offer.png", 5.7, 4.58, 2.3, 1.85)
    add_image(slide, "store-cards.png", 8.25, 4.58, 2.3, 1.85)
    add_image(slide, "email-launch.png", 10.8, 4.58, 1.9, 1.85)
    text_box(slide, 0.6, 6.55, 12.0, 0.32,
             "Bolt працює як платформа попиту: фокусує marketing inventory, Plus і free delivery на категорії та керує courier allocation у пікові lunch/dinner періоди.",
             8.5, MUTED, italic=True)
    footer(slide, 3)

    prs.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build()
