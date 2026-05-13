# Sử dụng Python 3.10 image gọn nhẹ
FROM python:3.10-slim

# Thiết lập thư mục làm việc
WORKDIR /app

# Copy file requirements và cài đặt
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ code vào container
COPY . .

# Chạy file run.py khi container khởi động
CMD ["python", "run.py"]