# Dockerfile for supplier-doc-automation
FROM python:3.11-slim
WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    unrar-free p7zip-full libarchive-tools nodejs npm \
    && rm -rf /var/lib/apt/lists/*

# Install pip deps FIRST (cache layer)
RUN pip install --no-cache-dir --default-timeout=120 -i https://mirrors.aliyun.com/pypi/simple/ python-docx requests openpyxl pyyaml pandas

# agently-cli（邮件监控/发送使用）
RUN npm install -g --registry=https://registry.npmmirror.com @tencent-qqmail/agently-cli

# Copy project files (invalidates cache on code change, but pip is already cached)
COPY scripts/ /app/scripts/
COPY agreement-automation/ /app/agreement-automation/
COPY email-database/ /app/email-database/
COPY data/ /app/data/

RUN mkdir -p /app/email-database/QA /app/email-database/output /app/output /app/data
ENV TZ=Asia/Shanghai

COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
CMD ["confirmer"]
