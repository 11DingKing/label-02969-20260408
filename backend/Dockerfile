# 多阶段构建：支持 ARM64 (Apple Silicon / AWS Graviton) 和 x86_64
FROM python:3.11-slim AS builder

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    git \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --user -r requirements.txt

# 运行时阶段
FROM python:3.11-slim

WORKDIR /app

COPY --from=builder /root/.local /root/.local
ENV PATH=/root/.local/bin:$PATH
ENV PYTHONIOENCODING=utf-8
ENV LANG=C.UTF-8

COPY word_formula_to_mathtype/ ./word_formula_to_mathtype/
COPY run_convert.py .
COPY scripts/ ./scripts/
RUN chmod +x scripts/entrypoint.sh

ENTRYPOINT ["bash", "scripts/entrypoint.sh"]
