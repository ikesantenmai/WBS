# Render などにデプロイするためのイメージ。
# .xls の読み書きに LibreOffice (soffice) を使うので、それも入れておく。
FROM python:3.12-slim

# libreoffice-calc: .xls <-> .xlsx の変換 / fonts-noto-cjk: 日本語の文字
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/* \
    && soffice --version

WORKDIR /app
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir ".[web]"

# LibreOffice は設定を HOME に書くので、書き込める場所にする
ENV HOME=/tmp
# LibreOffice の場所 (PATH に頼らず、アプリが直接見つけられるようにする)
ENV WBSGEN_SOFFICE=/usr/bin/soffice

# PORT があればそれで、0.0.0.0 で待ち受ける (Render が PORT を渡す)
CMD ["wbsgen", "serve"]
