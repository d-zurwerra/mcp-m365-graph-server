COPY requirements.txt .
RUN ls -la && echo "---- requirements.txt ----" && cat requirements.txt && echo "--------------------------"
RUN python -m pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt -v
