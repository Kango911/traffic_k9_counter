<!--
ЦВЕТОВАЯ ПАЛИТРА:
- Розовый: #FF69B4, #FF85C1, #FFB6C1
- Фиолетовый: #9B59B6, #BB8FCE, #D2B4DE
- Синий: #3498DB, #5DADE2, #85C1E9
- Градиенты: #FF69B4 → #9B59B6 → #3498DB
-->

<div align="center">

# 🚦 Traffic K9 Counter

### *Профессиональный счётчик транспорта на перекрёстке*

![Version](https://img.shields.io/badge/version-2.0.0-FF69B4?style=for-the-badge&labelColor=9B59B6)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-3498DB?style=for-the-badge&labelColor=9B59B6)
![License](https://img.shields.io/badge/license-MIT-FF69B4?style=for-the-badge&labelColor=9B59B6)
![Python](https://img.shields.io/badge/Python-3.8%2B-3498DB?style=for-the-badge&labelColor=9B59B6)

</div>

---

## ✨ Оглавление

1. [🌟 О программе](#-о-программе)
2. [🎯 Основные возможности](#-основные-возможности)
3. [🚀 Установка](#-установка)
4. [📖 Инструкция по использованию](#-инструкция-по-использованию)
5. [🎨 Главные "фишки"](#-главные-фишки)
6. [💾 Экспорт в Excel](#-экспорт-в-excel)
7. [❓ Часто задаваемые вопросы](#-часто-задаваемые-вопросы)
8. [📞 Контакты](#-контакты)

---

<div align="center">

## 🌟 О программе

</div>

**Traffic K9 Counter** — это мощное и интуитивно понятное приложение для ручного учёта транспортных средств на перекрёстке. Программа создана для специалистов по транспортной логистике, студентов и всех, кто занимается анализом дорожного движения.

<div align="center">

```mermaid
graph LR
    A[Въезды: N, S, E, W] --> B[Повороты: ➡️ ⬆️ ⬅️ 🔄]
    B --> C[10+ типов ТС]
    C --> D[📊 Excel-отчёт]
    
    style A fill:#FF69B4,stroke:#9B59B6,color:#fff
    style B fill:#9B59B6,stroke:#3498DB,color:#fff
    style C fill:#3498DB,stroke:#9B59B6,color:#fff
    style D fill:#FF69B4,stroke:#3498DB,color:#fff
```

</div>

---

<div align="center">

## 🎯 Основные возможности

</div>

<table align="center">
<tr>
<td width="33%" align="center">

### 🎛️ **Гибкие настройки**
Выбор любых въездов и выездов перед стартом

</td>
<td width="33%" align="center">

### 🚗 **10 типов ТС**
От легковых до трамваев + создание своих

</td>
<td width="33%" align="center">

### 🖱️ **Умные счётчики**
Левая кнопка +1, правая кнопка –1

</td>
</tr>
<tr>
<td width="33%" align="center">

### 💾 **Автосохранение**
Настройки и типы ТС запоминаются

</td>
<td width="33%" align="center">

### 📊 **Excel-отчёты**
Полная аналитика с долями поворотов

</td>
<td width="33%" align="center">

### 🎨 **Адаптивный интерфейс**
Изменяемый размер окна, всё под рукой

</td>
</tr>
</table>

---

<div align="center">

## 🚀 Установка

| Скачать для своей ОС |
|:---:|
| [![Windows](https://img.shields.io/badge/Windows-10%2F11-3498DB?style=for-the-badge&logo=windows&logoColor=white&labelColor=9B59B6)](https://github.com/Kango911/traffic_k9_counter/releases) |
| [![macOS](https://img.shields.io/badge/macOS-11.0%2B-FF69B4?style=for-the-badge&logo=apple&logoColor=white&labelColor=9B59B6)](https://github.com/Kango911/traffic_k9_counter/releases) |

</div>

### 📥 Пошаговая установка

<details>
<summary><b>🔷 Windows</b></summary>

1. Скачайте установщик `.exe` со [страницы релизов](https://github.com/Kango911/traffic_k9_counter/releases)
2. Запустите скачанный файл
3. Следуйте инструкциям установщика
4. Готово! Ярлык появится на рабочем столе

</details>

<details>
<summary><b>🍎 macOS</b></summary>

1. Скачайте файл `.dmg` со [страницы релизов](https://github.com/Kango911/traffic_k9_counter/releases)
2. Откройте скачанный образ
3. Перетащите `TrafficK9Counter.app` в папку `Applications`
4. Если система предупредит о неподтверждённом разработчике:
   - Нажмите **Отмена**
   - Откройте `System Preferences → Security & Privacy`
   - Нажмите **Всё равно открыть**

</details>

<details>
<summary><b>🐍 Из исходников (Python)</b></summary>

```bash
git clone https://github.com/Kango911/traffic_k9_counter.git
cd traffic_k9_counter
python -m venv venv
source venv/bin/activate  # или venv\Scripts\activate на Windows
pip install PySide6 openpyxl
python traffic_counter.py
```
</details>

---

<div align="center">

## 📖 Инструкция по использованию

</div>

### 🔷 Шаг 1: Выбор направлений

<div align="center">

| Въезды (откуда едут) | Выезды (куда поворачивают) |
|:---:|:---:|
| ⬆️ **N** (Север) | ➡️ **N** (Север) |
| ➡️ **E** (Восток) | ➡️ **E** (Восток) |
| ⬇️ **S** (Юг) | ⬇️ **S** (Юг) |
| ⬅️ **W** (Запад) | ⬅️ **W** (Запад) |

</div>

> 💡 **Пример**: Если перекрёсток имеет форму буквы «Т», выберите въезды «S», «E», «W», а выезды — только те, куда реально можно повернуть.

### 🔷 Шаг 2: Заполнение данных

В главном окне введите:
- 📍 **Название перекрёстка** (например, «Пр. Ленина / ул. Садовая»)
- 📅 **Дату** проведения учёта

### 🔷 Шаг 3: Настройка типов ТС (по желанию)

<div align="center">

![Типы ТС](https://img.shields.io/badge/🚚_Управление_типами_ТС-Нажмите_кнопку-FF69B4?style=for-the-badge)

</div>

- Нажмите **«Управление типами ТС»**
- Вы можете:
  - ➕ **Добавить** новый тип (например, «Велосипеды»)
  - ✏️ **Редактировать** существующие
  - 🗑️ **Удалить** ненужные
  - ☑️ **Включать/отключать** чекбоксы для скрытия типов

### 🔷 Шаг 4: Подсчёт транспорта

<div align="center">

**Порядок направлений:** `Направо → Прямо → Налево → Разворот`

</div>

| Действие | Результат |
|:---:|:---:|
| 🖱️ **Левая кнопка** | Увеличить счётчик на 1 |
| 🖱️ **Правая кнопка** | Уменьшить счётчик на 1 (не ниже 0) |
| ℹ️ **Кнопка `i`** | Показать описание типа ТС |

### 🔷 Шаг 5: Экспорт результатов

1. Нажмите **«Экспорт в Excel»**
2. Выберите имя файла и папку
3. Готово! Отчёт готов к анализу

---

<div align="center">

## 🎨 Главные «фишки»

</div>

<table>
<tr>
<td align="center" width="50%">

### 🎛️ **Динамические типы ТС**

</td>
<td align="center" width="50%">

### 💾 **Умное автосохранение**

</td>
</tr>
<tr>
<td valign="top">

- ➕ Добавляйте свои типы транспорта
- ✏️ Редактируйте названия и описания
- 🗑️ Удаляйте ненужные
- ☑️ Отключайте типы через чекбоксы
- 📥 Импорт/экспорт JSON-файлов

</td>
<td valign="top">

- 🔄 Все настройки сохраняются автоматически
- 📂 Выбор источника при запуске:
  - 💾 Автосохранённые настройки
  - 📄 Внешний JSON-файл
  - 🏭 Стандартный набор

</td>
</tr>
<tr>
<td align="center" colspan="2">

### 🎯 **Профессиональная аналитика в Excel**

</td>
</tr>
<tr>
<td colspan="2" valign="top">

- 📊 Полная таблица «направление → тип ТС»
- 🔢 Общее количество транспортных средств
- 🚌 Количество и доля ТС без общественного транспорта
- 📈 Доли поворотов для каждого въезда

</td>
</tr>
</table>

---

<div align="center">

## 💾 Экспорт в Excel

</div>

### 📋 Что содержит отчёт

<div align="center">

| Раздел | Описание |
|:---:|:---|
| **Шапка** | Название перекрёстка и дата учёта |
| **Основная таблица** | Матрица «Въезд → Поворот → Тип ТС» с количеством |
| **Общие итоги** | Сумма всех ТС |
| **Аналитика без транспорта** | Количество легковых и грузовых, их доля |
| **Доли поворотов** | Для каждого въезда — % направо, прямо, налево, разворот |

</div>

### 📊 Пример структуры отчёта

```
Перекрёсток: Пр. Ленина / ул. Садовая
Дата: 2026-06-11

          │ Легковые │ Автобусы │ Грузовые │ Итого
──────────┼──────────┼──────────┼──────────┼───────
Север↗️    │    12    │    3     │    5     │   20
Север⬆️    │    8     │    1     │    2     │   11
Север↖️    │    4     │    0     │    1     │    5
Север🔄    │    1     │    0     │    0     │    1
...
```

---

<div align="center">

## ❓ Часто задаваемые вопросы

</div>

<details>
<summary><b>❔ Можно ли изменить порядок направлений?</b></summary>

Нет, порядок **направо → прямо → налево → разворот** фиксированный, так как соответствует логике стандартного перекрёстка и единообразию отчётов.

</details>

<details>
<summary><b>❔ Что делать, если я случайно закрыл окно выбора направлений?</b></summary>

Закройте приложение и откройте заново — диалог появится снова при старте.

</details>

<details>
<summary><b>❔ Где хранятся мои добавленные типы ТС?</b></summary>

В файле `vehicle_types_auto.json` в папке с приложением. Вы также можете вручную экспортировать их через кнопку в окне управления типами.

</details>

<details>
<summary><b>❔ Можно ли экспортировать данные в старый Excel (XLS)?</b></summary>

Нет, поддерживается только современный формат XLSX (Excel 2007+).

</details>

<details>
<summary><b>❔ Программа бесплатная?</b></summary>

Да! Traffic K9 Counter распространяется под лицензией **MIT** — полностью бесплатно как для личного, так и для коммерческого использования.

</details>

---

<div align="center">

## 📞 Контакты

### По вопросам и предложениям

[![Telegram](https://img.shields.io/badge/Telegram-@Kango911-3498DB?style=for-the-badge&logo=telegram&logoColor=white&labelColor=9B59B6)](https://t.me/Kango911)

[![GitHub](https://img.shields.io/badge/GitHub-Kango911-FF69B4?style=for-the-badge&logo=github&logoColor=white&labelColor=9B59B6)](https://github.com/Kango911)

[![Issues](https://img.shields.io/badge/Сообщить_об_ошибке-Создать_Issue-9B59B6?style=for-the-badge&logo=github&logoColor=white&labelColor=FF69B4)](https://github.com/Kango911/traffic_k9_counter/issues)

</div>

---

<div align="center">

## ⭐ Поддержка проекта

Если вам понравилось приложение, поставьте звезду на GitHub!

[![Star](https://img.shields.io/badge/⭐_Поставить_звезду-На_GitHub-FF69B4?style=for-the-badge&logo=github&logoColor=white&labelColor=9B59B6)](https://github.com/Kango911/traffic_k9_counter)

**Traffic K9 Counter v2.0.0**  
*Лицензия MIT*

</div>

---

<div align="center">
  
**Сделано с ❤️ для точного учёта на дорогах**

</div>
