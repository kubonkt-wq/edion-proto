import io
import os
import json
from datetime import datetime
from typing import List
from fastapi import FastAPI
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

app = FastAPI()

# 日本語フォント登録
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))

# プロジェクト直下のルートディレクトリ取得
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

@app.get("/", response_class=HTMLResponse)
async def serve_home():
    html_path = os.path.join(BASE_DIR, "index.html")
    if not os.path.exists(html_path):
        html_path = "index.html"
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()

@app.get("/api/stores")
async def get_stores():
    json_path = os.path.join(BASE_DIR, "stores.json")
    if not os.path.exists(json_path):
        json_path = "stores.json"
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

class OrderItem(BaseModel):
    order_id: str
    order_date: str
    store_id: str = ""
    store_name: str
    branch: str = ""
    product_name: str
    price: int

class ExportPayload(BaseModel):
    orders: List[OrderItem]

@app.post("/api/export-excel")
async def export_excel(payload: ExportPayload):
    data = [item.dict() for item in payload.orders]
    df = pd.DataFrame(data)
    if df.empty:
        df = pd.DataFrame(columns=["order_id", "order_date", "store_id", "branch", "store_name", "product_name", "price"])
    
    df.rename(columns={
        "order_id": "発注番号",
        "order_date": "発注日時",
        "store_id": "販売店ID",
        "branch": "管轄支社",
        "store_name": "販売店名",
        "product_name": "商品名・型番",
        "price": "金額(円)"
    }, inplace=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="月次請求対象一覧")
    output.seek(0)

    filename = f"chunichi_billing_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/api/export-pdf")
async def export_pdf(payload: ExportPayload):
    output = io.BytesIO()
    p = canvas.Canvas(output, pagesize=A4)
    p.setFont("HeiseiMin-W3", 16)
    
    p.drawString(50, 800, "発注書（EDION 御中）")
    p.setFont("HeiseiMin-W3", 10)
    p.drawString(400, 800, f"発行日: {datetime.now().strftime('%Y年%m月%d日')}")
    p.drawString(400, 785, "発注元: 株式会社中日新聞社")
    p.line(50, 770, 550, 770)
    
    p.drawString(50, 750, "発注番号")
    p.drawString(130, 750, "販売店（配送先識別）")
    p.drawString(280, 750, "発注商品名")
    p.drawString(480, 750, "金額（税込）")
    p.line(50, 740, 550, 740)
    
    y = 720
    total = 0
    for item in payload.orders:
        p.drawString(50, y, item.order_id)
        display_store = f"[{item.store_id}] {item.store_name}" if item.store_id else item.store_name
        p.drawString(130, y, display_store[:12])
        p.drawString(280, y, item.product_name[:16])
        p.drawString(480, y, f"¥{item.price:,}")
        total += item.price
        y -= 25
        if y < 100:
            break
            
    p.line(50, y + 10, 550, y + 10)
    p.drawString(380, y - 10, f"合計発注金額: ¥{total:,}")
    
    p.showPage()
    p.save()
    output.seek(0)

    filename = f"edion_purchase_order_{datetime.now().strftime('%Y%m%d')}.pdf"
    return StreamingResponse(
        output,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )