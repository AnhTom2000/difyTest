# main.py
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
import uvicorn
from your_parser import analyze_docx_attachments

app = FastAPI(
    title="Dify DOCX Analyzer Plugin",
    description="一个用于解析DOCX附件中高危命令的Dify插件后端服务",
)

@app.post(
    "/analyze-docx", 
    summary="分析DOCX文件附件",
    description="上传一个.docx文件，服务将提取其中的附件并解析潜在的风险命令。"
)
async def analyze_docx_endpoint(file: UploadFile = File(..., description="用户上传的.docx文件")):
    
    # 验证文件类型
    if not file.filename.endswith('.docx'):
        return JSONResponse(
            status_code=400,
            content={"error": "文件格式错误，请上传.docx文件。"}
        )

    try:
        # 读取上传文件的二进制内容
        docx_binary_data = await file.read()
        
        # 调用我们的核心解析逻辑
        results = analyze_docx_attachments(docx_binary_data)
        
        if "error" in results:
             return JSONResponse(status_code=500, content=results)

        # 成功后返回JSON格式的结果
        return JSONResponse(status_code=200, content=results)

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": f"处理文件时发生意外错误: {str(e)}"}
        )

# 方便本地测试
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)