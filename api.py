import os
import asyncio
import traceback
from concurrent.futures import ThreadPoolExecutor
from fastapi.middleware.cors import CORSMiddleware
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Body, Depends
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from fastapi.responses import StreamingResponse
from module import extraction


# Define global environment vairables
BACKEND_API_SECRET_KEY=os.getenv("BACKEND_API_SECRET_KEY")


app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=[os.getenv("ALLOWED_CORS")],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

thread_executors = ThreadPoolExecutor(max_workers=int(os.getenv("THREAD_NUMBERS")))


@app.post("/extract-data")
async def extract_data(file_bytes: UploadFile = File(...),
                       credentials: HTTPAuthorizationCredentials = Depends(HTTPBearer())):
    """
    Function as an API endpoint to response the user's chat from front-end
    """
    async_loop = asyncio.get_running_loop()

    # Check credentials
    token_bearer = credentials.credentials
    if str(token_bearer) != BACKEND_API_SECRET_KEY:
        raise HTTPException(status_code=403, detail="Unauthorized request")

    try:
        file_bytes = await file_bytes.read()
        buffer_output = await async_loop.run_in_executor(
            thread_executors,
            extraction,
            file_bytes
        )

        return StreamingResponse(
            buffer_output
        )
    except Exception:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail="Failed processing extraction")