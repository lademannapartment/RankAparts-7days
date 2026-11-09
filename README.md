# RaksAparts-7days

Скан позиций объектов на Booking.com на 7 ночей. Результаты пишутся в лист "Rank AP7d".

## Запуск

```
pip install -r requirements.txt
python -m playwright install chromium
python RankAP7d.py
```

- `GOOGLE_CREDS` — JSON сервисного аккаунта в переменной окружения (или файл `silken-glyph-443313-i4-cdf14b36288a.json` рядом со скриптом).
- `HEADLESS=1` — запуск без окна браузера (в GitHub Actions уже задан).
