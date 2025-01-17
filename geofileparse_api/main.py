# es_api/main.py
import sys
import os

sys.path.append(os.getcwd())
from fastapi import FastAPI
from geofileparse_api.routers import keyanexport_router  # 确保路径基于项目根

app = FastAPI()

# 注册路由
app.include_router(keyanexport_router.router, prefix="/esapi", tags=["keyanexport"])

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("es_api.main:app", host="0.0.0.0", port=8080, reload=True)