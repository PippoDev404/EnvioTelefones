"""
API de Transcrição de Áudio usando FastAPI + Faster-Whisper
"""
import os
import tempfile
import asyncio
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from faster_whisper import WhisperModel

# Configurações
MODEL_SIZE = "medium"  # tiny, base, small, medium, large-v2
LANGUAGE = "pt"
DEVICE = "auto"  # auto, cpu, cuda

# Inicializa o modelo globalmente (carrega uma vez)
print(f"Carregando modelo Whisper {MODEL_SIZE}...")
whisper_model = WhisperModel(
    MODEL_SIZE,
    device=DEVICE,
    compute_type="int8"  # int8, float16, float32
)
print("Modelo carregado com sucesso!")

app = FastAPI(title="API de Transcrição")

# CORS - permitir requisições do seu frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def transcrever_audio(caminho_arquivo: str) -> dict:
    """
    Transcreve um arquivo de áudio usando Faster-Whisper.
    """
    try:
        segments, info = whisper_model.transcribe(
            caminho_arquivo,
            language=LANGUAGE,
            beam_size=5,
            vad_filter=True,  # Filtro de detecção de voz
            vad_parameters=dict(min_silence_duration_ms=500)
        )

        # Converte generator em lista e concatena
        texto_completo = ""
        chunks = []

        for segment in segments:
            texto_completo += segment.text.strip() + " "
            chunks.append({
                "text": segment.text.strip(),
                "start": segment.start,
                "end": segment.end
            })

        return {
            "text": texto_completo.strip(),
            "chunks": chunks,
            "language": info.language,
            "language_probability": info.language_probability
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro na transcrição: {str(e)}")


@app.get("/")
async def root():
    return {"status": "ok", "message": "API de Transcrição ativa", "model": MODEL_SIZE}


@app.get("/health")
async def health():
    return {"status": "healthy", "model": MODEL_SIZE}


@app.post("/transcrever")
async def transcrever(arquivo: UploadFile = File(...)):
    """
    Endpoint para transcrever um arquivo de áudio.
    """
    # Valida tipo do arquivo
    if not arquivo.content_type or not arquivo.content_type.startswith("audio/"):
        raise HTTPException(status_code=400, detail="Arquivo deve ser um áudio válido")

    # Salva arquivo temporário
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(arquivo.filename).suffix) as tmp:
        conteudo = await arquivo.read()
        tmp.write(conteudo)
        tmp_path = tmp.name

    try:
        resultado = transcrever_audio(tmp_path)
        
        return JSONResponse(content={
            "success": True,
            "nomeEntrevista": arquivo.filename,
            "transcricao": resultado["text"],
            "chunks": resultado["chunks"],
            "language": resultado["language"]
        })

    finally:
        # Limpa arquivo temporário
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.post("/transcrever-multiplos")
async def transcrever_multiplos(arquivos: List[UploadFile] = File(...)):
    """
    Endpoint para transcrever múltiplos arquivos de áudio.
    """
    if not arquivos or len(arquivos) == 0:
        raise HTTPException(status_code=400, detail="Nenhum arquivo enviado")

    if len(arquivos) > 10:
        raise HTTPException(status_code=400, detail="Máximo de 10 arquivos por vez")

    resultados = []

    for arquivo in arquivos:
        # Valida tipo do arquivo
        if not arquivo.content_type or not arquivo.content_type.startswith("audio/"):
            resultados.append({
                "nomeEntrevista": arquivo.filename,
                "transcricao": "",
                "erro": "Arquivo inválido"
            })
            continue

        # Salva arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(arquivo.filename).suffix) as tmp:
            conteudo = await arquivo.read()
            tmp.write(conteudo)
            tmp_path = tmp.name

        try:
            resultado = transcrever_audio(tmp_path)
            resultados.append({
                "nomeEntrevista": arquivo.filename,
                "transcricao": resultado["text"],
                "chunks": resultado["chunks"],
                "language": resultado["language"]
            })
        except Exception as e:
            resultados.append({
                "nomeEntrevista": arquivo.filename,
                "transcricao": "",
                "erro": str(e)
            })
        finally:
            # Limpa arquivo temporário
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    return JSONResponse(content={
        "success": True,
        "resultados": resultados
    })


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)