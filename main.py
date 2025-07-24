#from fastapi import FastAPI
from metabase_extract import MetabaseDashboardExtract
import os
import google.generativeai as genai

#app = FastAPI()

#@app.post("/exportar-informe")
def exportar():
    email = os.getenv("METABASE_EMAIL")
    password = os.getenv("METABASE_PASSWORD")
    base_url = os.getenv("METABASE_BASE_URL")
    api_key = os.getenv("GEMINI_API_KEY")

    if not email or not password or not api_key or not base_url:
        return {"error": "Faltan variables de entorno"}

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.0-flash')

    try:
        extractor = MetabaseDashboardExtract(email, password, base_url, model=model)
        extractor.run()
        return {"status": "Informe generado correctamente"}
    except Exception as e:
        return {"error": str(e)}




if __name__ == "__main__":
    result = exportar()
    print(result)
