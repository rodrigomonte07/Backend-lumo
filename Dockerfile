FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-liberation \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Fontes do template (Eastman Alternate, Neulis Neue, Nourd, Lato) — sem elas o LibreOffice troca por
# outra fonte e o PDF sai com letras espaçadas e textos quebrando. Veja fonts/LEIAME.txt
RUN mkdir -p /usr/local/share/fonts/rede-lumo && cp fonts/* /usr/local/share/fonts/rede-lumo/ && fc-cache -f

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
