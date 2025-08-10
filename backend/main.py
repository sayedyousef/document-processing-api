"""
Fixed Backend main.py - Creates actual output files
Save this as: backend/main.py
"""

from fastapi import FastAPI, File, UploadFile, BackgroundTasks, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import uuid
from pathlib import Path
from typing import List
import asyncio
import zipfile
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Document Processing API")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create necessary directories
BASE_DIR = Path(__file__).parent
TEMP_DIR = BASE_DIR / "temp"
OUTPUT_DIR = BASE_DIR / "output"
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# In-memory job tracking
jobs = {}

@app.post("/api/process")
async def process_documents(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    processor_type: str = Form("scan_verify")
):
    """Process documents with specified processor"""
    job_id = str(uuid.uuid4())
    
    logger.info(f"New job {job_id}: Processing {len(files)} files with {processor_type}")
    
    # Initialize job
    jobs[job_id] = {
        "status": "processing",
        "total": len(files),
        "completed": 0,
        "results": [],
        "processor": processor_type
    }
    
    # Save uploaded files
    temp_dir = TEMP_DIR / job_id
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    file_paths = []
    for file in files:
        file_path = temp_dir / file.filename
        content = await file.read()
        file_path.write_bytes(content)
        file_paths.append(file_path)
        logger.info(f"Saved file: {file_path}")
    
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
        "message": f"Processing {len(files)} documents with {processor_type}"
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
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if job["status"] != "completed":
        raise HTTPException(status_code=400, detail="Job not completed")
    
    results = job["results"]
    
    # Filter out failed results
    successful_results = [r for r in results if "error" not in r]
    
    if not successful_results:
        raise HTTPException(status_code=404, detail="No successful results to download")
    
    if len(successful_results) == 1:
        # Single file - return directly
        file_path = Path(successful_results[0]["path"])
        if not file_path.exists():
            raise HTTPException(status_code=404, detail=f"File not found: {file_path}")
        return FileResponse(
            path=str(file_path),
            filename=successful_results[0]["output_filename"],
            media_type="application/octet-stream"
        )
    else:
        # Multiple files - create ZIP
        zip_path = TEMP_DIR / job_id / f"results_{job_id}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for result in successful_results:
                file_path = Path(result["path"])
                if file_path.exists():
                    zipf.write(file_path, result["output_filename"])
        
        if not zip_path.exists():
            raise HTTPException(status_code=404, detail="Failed to create ZIP file")
            
        return FileResponse(
            path=str(zip_path),
            filename=f"results_{job_id}.zip",
            media_type="application/zip"
        )

@app.get("/api/download/{job_id}/{index}")
async def download_single_result(job_id: str, index: int):
    """Download a single result file"""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if index >= len(job["results"]):
        raise HTTPException(status_code=404, detail="Invalid file index")
    
    result = job["results"][index]
    if "error" in result:
        raise HTTPException(status_code=400, detail=f"File processing failed: {result['error']}")
    
    file_path = Path(result["path"])
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"File not found: {file_path}")
        
    return FileResponse(
        path=str(file_path),
        filename=result["output_filename"],
        media_type="application/octet-stream"
    )

async def process_job(job_id: str, file_paths: List[Path], processor_type: str):
    """Background job processor"""
    logger.info(f"Starting background processing for job {job_id}")
    
    # Import processors here to avoid circular imports
    if processor_type == "word_to_html":
        from processors.word_to_html_processor import WordToHtmlProcessor
        processor = WordToHtmlProcessor()
    elif processor_type == "scan_verify":
        from processors.scan_verify_processor import ScanVerifyProcessor
        processor = ScanVerifyProcessor()
    else:
        logger.error(f"Unknown processor type: {processor_type}")
        jobs[job_id]["status"] = "failed"
        return
    
    output_dir = OUTPUT_DIR / job_id
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for i, file_path in enumerate(file_paths):
        try:
            logger.info(f"Processing file {i+1}/{len(file_paths)}: {file_path.name}")
            
            # Process the file
            result = await processor.process(file_path, output_dir)
            
            # Add index for individual downloads
            result["index"] = i
            result["job_id"] = job_id
            
            jobs[job_id]["results"].append(result)
            jobs[job_id]["completed"] += 1
            
            logger.info(f"Successfully processed: {file_path.name}")
            
        except Exception as e:
            logger.error(f"Error processing {file_path.name}: {str(e)}")
            jobs[job_id]["results"].append({
                "filename": file_path.name,
                "error": str(e),
                "index": i
            })
            jobs[job_id]["completed"] += 1
    
    jobs[job_id]["status"] = "completed"
    logger.info(f"Job {job_id} completed: {jobs[job_id]['completed']} files processed")

@app.get("/")
async def root():
    """Health check endpoint"""
    return {"status": "running", "message": "Document Processing API"}

@app.get("/api/jobs")
async def list_jobs():
    """List all jobs (for debugging)"""
    return {
        "total_jobs": len(jobs),
        "jobs": [
            {
                "job_id": job_id,
                "status": job_data["status"],
                "processor": job_data["processor"],
                "progress": f"{job_data['completed']}/{job_data['total']}"
            }
            for job_id, job_data in jobs.items()
        ]
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)