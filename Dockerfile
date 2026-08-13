# Общий образ для Telegram-бота (bot.py) и MAX-бота (max.py).

FROM python:3.11-slim

# Системные зависимости:
# - libjpeg/zlib нужны Pillow для работы с изображениями
# - libxml2/libxslt нужны lxml
# - build-essential нужен для сборки некоторых колёс "с нуля"
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        libjpeg62-turbo-dev \
        zlib1g-dev \
        libxml2-dev \
        libxslt1-dev \
        ca-certificates \
        wget \
        antiword \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Сначала зависимости — чтобы Docker кэшировал этот слой
# и не переустанавливал пакеты при каждом изменении кода
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем код проекта
COPY core ./core
COPY bot.py max.py ./
# Если у вас есть свои шаблоны сообщений/картинок — раскомментируйте:
# COPY messages ./messages
# COPY images ./images

# Логи и изображения бот создаёт сам при старте (misc.configure_logging,
# IMAGES_DIR.mkdir), но на всякий случай подготовим директории заранее
RUN mkdir -p logs images messages

# Конкретная команда запуска (bot.py или max.py) задаётся
# в docker-compose.yml для каждого сервиса отдельно.
CMD ["python", "bot.py"]
