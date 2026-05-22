import base64
import io
import os

import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from openai import OpenAI
from PIL import Image
from rembg import remove

load_dotenv()

app = FastAPI(title="AI Image Generator with Transparent Background")
app.mount("/static", StaticFiles(directory="static"), name="static")

client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))


@app.get("/", response_class=HTMLResponse)
async def index():
    with open("static/index.html", encoding="utf-8") as f:
        return f.read()


@app.post("/generate")
async def generate_image(
    prompt: str = Form(...),
    size: str = Form("1024x1024"),
    remove_bg: str = Form("true"),
):
    if not os.environ.get("OPENAI_API_KEY"):
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY が設定されていません")

    try:
        response = client.images.generate(
            model="gpt-image-1",
            prompt=prompt,
            n=1,
            size=size,
        )

        image_data_b64 = response.data[0].b64_json
        if image_data_b64:
            image_bytes = base64.b64decode(image_data_b64)
        else:
            image_url = response.data[0].url
            async with httpx.AsyncClient() as http_client:
                img_response = await http_client.get(image_url, timeout=30)
                image_bytes = img_response.content

        if remove_bg == "true":
            input_image = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
            output_image = remove(input_image)
            output_buffer = io.BytesIO()
            output_image.save(output_buffer, format="PNG")
            output_buffer.seek(0)
            result_bytes = output_buffer.read()
            media_type = "image/png"
        else:
            result_bytes = image_bytes
            media_type = "image/png"

        return Response(
            content=result_bytes,
            media_type=media_type,
            headers={"Content-Disposition": "attachment; filename=generated.png"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
