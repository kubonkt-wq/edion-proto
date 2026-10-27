import io
from datetime import datetime
from typing import List
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

app = FastAPI()

# 日本語フォント登録（標準の平成明朝）
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))

class OrderItem(BaseModel):
    order_id: str
    order_date: str
    store_name: str
    product_name: str
    price: int

class ExportPayload(BaseModel):
    orders: List[OrderItem]

@app.post("/api/export-excel")
async def export_excel(payload: ExportPayload):
    """新聞社が月次請求書に転記するためのExcel生成"""
    data = [item.dict() for item in payload.orders]
    df = pd.DataFrame(data)
    if df.empty:
        df = pd.DataFrame(columns=["order_id", "order_date", "store_name", "product_name", "price"])
    
    df.rename(columns={
        "order_id": "発注番号",
        "order_date": "発注日時",
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
    """中日新聞社からEDIONへ送付する発注書PDF生成"""
    output = io.BytesIO()
    p = canvas.Canvas(output, pagesize=A4)
    p.setFont("HeiseiMin-W3", 16)
    
    # ヘッダー情報
    p.drawString(50, 800, "発注書（EDION 御中）")
    p.setFont("HeiseiMin-W3", 10)
    p.drawString(400, 800, f"発行日: {datetime.now().strftime('%Y年%m月%d日')}")
    p.drawString(400, 785, "発注元: 株式会社中日新聞社")
    
    # 罫線
    p.line(50, 770, 550, 770)
    
    # テーブル見出し
    p.drawString(50, 750, "発注番号")
    p.drawString(130, 750, "販売店（配送先識別）")
    p.drawString(280, 750, "発注商品名")
    p.drawString(480, 750, "金額（税込）")
    p.line(50, 740, 550, 740)
    
    # 注文明細の出力
    y = 720
    total = 0
    for item in payload.orders:
        p.drawString(50, y, item.order_id)
        p.drawString(130, y, item.store_name[:12])
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