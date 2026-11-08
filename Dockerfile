FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV SLT_PM_HOST=0.0.0.0 SLT_PM_PORT=5000 SLT_PM_DB=sqlite:////data/slt_pm.db
VOLUME /data
EXPOSE 5000
CMD ["python", "app.py"]
