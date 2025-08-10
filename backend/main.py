from fastapi import FastAPI, File, UploadFile, BackgroundTasks, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import uuid
from pathlib import Path
from typing import List
import asyncio

app = FastAPI(title="Document Processing API")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job tracking
jobs = {}

@app.post("/api/process")
async def process_documents(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    processor_type: str = Form("scan_verify")  # scan_verify or word_to_html
):
    """Process documents with specified processor"""
    job_id = str(uuid.uuid4())
    
    # Initialize job
    jobs[job_id] = {
        "status": "processing",
        "total": len(files),
        "completed": 0,
        "results": [],
        "processor": processor_type
    }
    
    # Save files and start processing
    temp_dir = Path(f"backend/temp/{job_id}")
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    file_paths = []
    for file in files:
        file_path = temp_dir / file.filename
        content = await file.read()
        file_path.write_bytes(content)
        file_paths.append(file_path)
    
    # Queue background processing
    background_tasks.add_task(
        process_job, 
        job_id, 
        file_paths, 
        processor_type
    )
    
    return {
        "job_id": job_id,
        "status": "processing",
        "message": f"Processing {len(files)} documents"
    }

@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    """Check processing status"""
    if job_id not in jobs:
        return JSONResponse(status_code=404, content={"error": "Job not found"})
    
    job = jobs[job_id]
    return {
        "job_id": job_id,
        "status": job["status"],
        "progress": f"{job['completed']}/{job['total']}",
        "processor": job["processor"],
        "results": job.get("results", [])
    }

@app.get("/api/download/{job_id}")
async def download_results(job_id: str):
    """Download processed results"""
    if job_id not in jobs:
        return JSONResponse(status_code=404, content={"error": "Job not found"})
    
    job = jobs[job_id]
    if job["status"] != "completed":
        return JSONResponse(status_code=400, content={"error": "Job not completed"})
    
    results = job["results"]
    if len(results) == 1:
        return FileResponse(results[0]["path"])
    else:
        # For multiple files, create a ZIP
        import zipfile
        zip_path = Path(f"backend/temp/{job_id}/results.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for result in results:
                zipf.write(result["path"], result["filename"])
        return FileResponse(zip_path, filename=f"results_{job_id}.zip")

async def process_job(job_id: str, file_paths: List[Path], processor_type: str):
    """Background job processor"""
    from processors.processor_factory import get_processor
    
    processor = get_processor(processor_type)
    
    for i, file_path in enumerate(file_paths):
        try:
            result = await processor.process(file_path)
            jobs[job_id]["results"].append(result)
            jobs[job_id]["completed"] += 1
        except Exception as e:
            print(f"Error processing {file_path}: {e}")
            jobs[job_id]["results"].append({
                "filename": file_path.name,
                "error": str(e)
            })
    
    jobs[job_id]["status"] = "completed"

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)