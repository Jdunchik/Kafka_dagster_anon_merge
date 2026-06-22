## 1. Правильная постановка задачи

- Entity Resolution / Record Linkage / Product Matching.
- Разница между:
  - `same entity`
  - `different entity`
  - `uncertain / needs review`
- Почему мэтчинг - это не просто похожесть строк.
- Почему для бизнеса false positive часто хуже false negative.

## 2. Метрики качества

- Golden dataset: пары значений с ручной разметкой `match / non-match`.
- Precision.
- Recall.
- F1.
- Confusion matrix.
- Отдельный анализ false positives и false negatives.
- Подбор порогов по метрикам, а не на глаз.

## 3. Нормализация строк

- Регистр, пробелы, пунктуация.
- `ё` / `е`.
- Кириллица / латиница: `Fairy` / `Фейри`, `Ariel` / `Ариэль`.
- Опечатки и сокращения.
- Удаление мусорных токенов: `арт`, `новинка`, `акция`, `штрихкод`, `код`.
- Словари синонимов и альтернативных написаний.

## 4. Извлечение атрибутов товара

- Бренд.
- Производитель.
- Категория.
- Тип товара.
- Объем / вес / размер.
- Количество в упаковке.
- Аромат / вкус / цвет / вариант.
- Формат упаковки.

Важно: разные объемы, вкусы, ароматы и цвета обычно нельзя склеивать автоматически.

## 5. Единицы измерения и упаковки

- `500мл`, `0.5л`, `500 ml` -> один формат.
- `1кг`, `1000г` -> один формат.
- `2x500мл`, `2*500мл`, `2 шт по 500 мл`.
- `10шт`, `20шт`.
- Наборы, мультипаки, промоупаковки.

## 6. Blocking / Candidate Generation

- Не сравнивать каждую строку с каждой.
- Сначала ограничивать кандидатов по:
  - штрихкоду
  - бренду
  - категории
  - размеру
  - ключевым токенам
- Изучить blocking в record linkage.

## 7. String Similarity

- Levenshtein distance.
- Jaro / Jaro-Winkler.
- Token sort ratio.
- Token set ratio.
- Partial ratio.
- Character n-grams.
- TF-IDF по char n-grams.

Библиотеки:

- RapidFuzz: https://rapidfuzz.github.io/RapidFuzz/
- scikit-learn `TfidfVectorizer`: https://scikit-learn.org/stable/modules/generated/sklearn.feature_extraction.text.TfidfVectorizer.html

## 8. Probabilistic / ML Entity Resolution

- Почему одного similarity score недостаточно.
- Комбинация признаков: бренд, категория, размер, токены, строковая похожесть.
- Confidence score.
- Explainable match: почему пара смэтчилась.
- Active learning / ручная разметка спорных пар.

Библиотеки:

- Dedupe: https://docs.dedupe.io/
- Splink: https://moj-analytical-services.github.io/splink/
- Python Record Linkage Toolkit: https://recordlinkage.readthedocs.io/

## 9. Semantic Similarity

- Sentence embeddings.
- Cosine similarity.
- Sentence Transformers.

Смотреть осторожно: embeddings могут считать похожими товары с разными SKU-атрибутами.

Источник:

- Sentence Transformers STS: https://sbert.net/docs/sentence_transformer/usage/semantic_textual_similarity.html

## 10. Кластеризация результатов

- Pairwise matching.
- Группировка в кластеры.
- Риск цепочек: `A ~ B`, `B ~ C`, но `A != C`.
- Выбор canonical ID.
- Выбор canonical name.
- Хранение всех исходных вариантов.

## 11. Human-in-the-loop

- Очередь спорных пар.
- Ручное подтверждение match.
- Ручное отклонение match.
- Хранение negative matches.
- История решений.
- Повторное использование ручной разметки.

## 12. Что должен уметь итоговый модуль

- Возвращать `match / non-match / review`.
- Возвращать score.
- Возвращать причины решения.
- Поддерживать ручные overrides.
- Иметь тестовый набор размеченных пар.
- Показывать precision / recall на тестовом наборе.
- Не склеивать товары с разными критичными атрибутами.
- Работать отдельно для товаров, брендов, компаний и магазинов.

