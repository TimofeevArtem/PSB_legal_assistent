# PSB_legal_assistent

# Chunk creation pipeline

Модуль `chunk_create.py` подготавливает базу знаний из ZIP-архива и возвращает готовые чанки для дальнейшего использования в RAG-системе.

## Основная функция

```python
from chunk_create import create_chunks

chunks = create_chunks(
    data_dir="C:/path/to/data",
    chunk_size=450,
    overlap=70,
    min_chunk_size=50,
)
```

## Как работает

Функция `create_chunks`:

1. принимает путь к папке `data_dir`;
2. ищет в этой папке архив `knowledge-base.zip`;
3. распаковывает архив в папку `knowledge-base` внутри `data_dir`;
4. читает поддерживаемые документы (`.txt`, `.md`, `.docx`);
5. разбивает текст по юридическим разделам и абзацам;
6. объединяет слишком маленькие чанки с соседями;
7. добавляет метаданные: документ, раздел, контрагент, дата редакции, статус действия и др.;
8. возвращает список словарей с готовыми чанками.

## Формат чанка

Каждый чанк представляет собой словарь со структурой вида:

```python
{
    "document_id": "...",
    "chunk_id": "...",
    "parent_id": "...",
    "document": "Название документа",
    "section": "1. Общие положения",
    "section_title": "Общие положения",
    "section_path": ["1"],
    "text": "Текст чанка",
    "context": "Контекст чанка",
    "embedding_text": "Текст для эмбеддингов",
    "token_count": 120,
    "counterparty": "Альфа-Трейд",
    "revision_date": "2024-01-01",
    "validity_status": "действует",
    "supplementary_agreement_for": "...",
}
```

## Переменная окружения

Если `data_dir` не передан явно, модуль пытается взять значение из `DATA_DIR`.

## Требования

- Python 3.10+
- `python-docx`
- `tiktoken`

## Запуск как скрипта

```bash
python chunk_create.py
```

При запуске через файл будет вызвана `create_chunks()` с `DATA_DIR` из окружения или с аргументами по умолчанию.
