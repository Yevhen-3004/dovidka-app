# Генератор довідки — Валютний контроль

Веб-додаток для генерації довідок про залишки іноземної валюти на рахунках ФОП.

---

## 🚀 Деплой на Render.com (безкоштовно)

### Крок 1 — Завантаж код на GitHub

1. Зайди на https://github.com та створи безкоштовний акаунт
2. Натисни "New repository" → назва: dovidka-app → "Create repository"
3. Натисни "uploading an existing file"
4. Перетягни всі файли (app.py, requirements.txt, render.yaml, build.sh)
5. Натисни "Commit changes"

### Крок 2 — Підключи до Render

1. Зайди на https://render.com → зареєструйся через GitHub
2. Натисни "New +" → "Web Service"
3. Обери репозиторій dovidka-app
4. Налаштування:
   - Runtime: Python 3
   - Build Command: bash build.sh
   - Start Command: gunicorn app:app
   - Instance Type: Free
5. Натисни "Create Web Service"

### Крок 3 — Готово!

Через 2-3 хвилини отримаєш посилання:
https://dovidka-app.onrender.com

Ділись цим посиланням з колегами — воно відкривається з будь-якого пристрою.

---

## ⚠️ Безкоштовний план Render

- Сервер засинає після 15 хв неактивності
- Перший запит після сну ~30 сек (далі миттєво)
- Платний план $7/міс — без обмежень

---

## 💻 Локальний запуск

pip install flask reportlab
python app.py

Відкрий: http://localhost:5000
