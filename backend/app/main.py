from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.config import settings
from app.api import auth, equipment, projects, systems, publish

app = FastAPI(
    title="AV BoM Tool",
    description="Audiovisual Equipment List & Bill of Materials Web Application",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[settings.FRONTEND_URL, "http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(auth.router)
app.include_router(equipment.router)
app.include_router(projects.router)
app.include_router(systems.router)
app.include_router(publish.router)


@app.get("/health")
async def health():
    return {"status": "ok"}
