from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import tempfile
import os
from excel_process import process_excel  # 导入你已有的处理函数

# 初始化FastAPI应用
app = FastAPI(title="学生宿舍分类工具")

# 配置前端模板（用于显示上传页面）
templates = Jinja2Templates(directory="templates")

# 模型路径（你的模型文件）
MODEL_PATH = "K_model.pkl"


# 1. 显示上传页面（GET请求）
@app.get("/", response_class=HTMLResponse)
async def show_upload_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# 2. 处理文件上传并返回结果（POST请求）
@app.post("/process")
async def process_file(file: UploadFile = File(...)):
    # 检查文件格式是否为Excel
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传.xlsx格式的Excel文件")

    # 创建临时文件存储上传的Excel
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
        temp_input.write(await file.read())
        input_path = temp_input.name

    # 创建临时文件存储处理后的结果
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_output:
        output_path = temp_output.name

    try:
        # 调用你的处理函数（传入输入路径、输出路径、模型路径）
        process_excel(
            input_excel_path=input_path,
            output_excel_path=output_path,
            model_path=MODEL_PATH
        )
    except Exception as e:
        # 出错时清理临时文件，避免占用空间
        os.unlink(input_path)
        os.unlink(output_path)
        raise HTTPException(status_code=500, detail=f"处理失败：{str(e)}")

    # 返回处理后的文件，并设置下载后自动删除临时文件
    return FileResponse(
        output_path,
        filename=f"分类结果_{file.filename}",  # 下载的文件名
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=lambda: [os.unlink(input_path), os.unlink(output_path)]  # 清理临时文件
    )


# 本地运行入口
if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
