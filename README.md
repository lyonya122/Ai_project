# AI Presentation Generator

Учебный проект: веб-приложение для генерации презентаций с помощью ИИ.
- веб-интерфейс для ввода идеи презентации
- 3 агента по ТЗ:
  - **planner / агент-структуризатор**
  - **writer / агент-копирайтер**
  - **designer / агент-дизайнер**
- координация агентов через **LangGraph**
- базовый **RAG** через ChromaDB + sentence-transformers
- экспорт итоговой презентации в **.pptx**
- Python 3.11-friendly конфигурация
- OpenAI-only конфигурация без Anthropic


## Быстрый запуск
### 1. Backend
```bash
cd backend
py -3.11 -m venv .venv
.venv\Scripts ctivate
pip install -r requirements.txt
copy .env.example .env
```
Запуск:
```bash
uvicorn app.main:app --reload
```

### 2. Frontend
```bash
cd frontend
npm install
npm run dev
```

1. Запусти backend.
2. Запусти frontend.
3. При желании загрузи PDF/TXT как базу знаний.
4. Нажми **Сгенерировать**.
5. Скачается `.pptx`.


