import os
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from docx import Document as docx_document
from docx.shared import Pt, RGBColor
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
from PIL import Image # Para abrir imagenes (Pillow)
from google.api_core.exceptions import ResourceExhausted
from docx2pdf import convert
from spire.doc import *
from spire.doc.common import *
import subprocess

class MetabaseDashboardExtract:
    # Constructor de clase (se ejecuta automaticamente al crear un objeto)
    def __init__(self, email, password, base_url, output_dir="output"):
        self.email = email
        self.password = password
        self.base_url = base_url
        self.output_dir = output_dir

        # Comportamiento de Chrome antes de abrirlo
        chrome_options = Options()
        # Sin interfaz grafica
        chrome_options.add_argument("--headless=new")
        # Aceleracion por hardware de la GPU (Deshabilitado)
        chrome_options.add_argument("--disable-gpu")
        # Como esta en headless, le damos una resolucion para que renderice
        chrome_options.add_argument("--window-size=1920,1080")

        prefs = {
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0
        }
        chrome_options.add_experimental_option("prefs", prefs)

        self.driver = webdriver.Chrome(options=chrome_options)
        os.makedirs(self.output_dir, exist_ok=True)

    def login(self):
        login_url = f"{self.base_url}/auth/login"
        self.driver.get(login_url)

        try:
            username_field = WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.NAME, "username"))
            )
            password_field = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "password"))
            )
            username_field.send_keys(self.email)
            password_field.send_keys(self.password)

            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
            )
            login_button.click()

            WebDriverWait(self.driver, 30).until(
                EC.url_contains("/")
            )
        except TimeoutException as e:
            print(f"Error en login: {e}")
            raise

    def capture_dashboard(self, dashboard_url, municipio, tab):

        print(f"capturando dashboard para {municipio} - {tab}")

        self.driver.get(dashboard_url)
        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="dashcard-container"]')))
        time.sleep(10)

        
        cards = self.driver.find_elements(By.CSS_SELECTOR, "div.elbzci2t1.emotion-7apf2f1.e1isodme0")
        viz_cards = self.driver.find_elements(By.CSS_SELECTOR, '[data-testid="visualization-root"]')
        all_cards = cards + viz_cards

        if not all_cards:
            print("No se encontraron gráficos.")
            return


        for i, card in enumerate(all_cards):
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card)
                time.sleep(2)
                WebDriverWait(self.driver, 10).until(EC.visibility_of(card))

                title = ""
                for selector in ["[data-testid='legend-caption-title']", "[data-testid='scalar-title']", ".Card-title", "h3", "h4"]:
                    try:
                        title_element = card.find_element(By.CSS_SELECTOR, selector)
                        title = title_element.text.strip()
                        if title:
                            break
                    except:
                        continue

                if not title:
                    title = f"grafico_{i+1}"

                clean_title = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", title)[:80]
                folder_path = os.path.join(self.output_dir, municipio, tab)
                os.makedirs(folder_path, exist_ok=True)

                filename = f"{clean_title}.png"
                filepath = os.path.join(folder_path, filename)

                card.screenshot(filepath)
                print(f"Capturado: {filepath}")

            except Exception as e:
                print(f"Error capturando gráfico {i+1}: {e}")
                continue
    
    @staticmethod
    def get_image_description_from_gemini(image_paths, prompt):
        pil_images = []
        try:
            for path in image_paths:
                img = Image.open(path)
                pil_images.append(img)
        except Exception as e:
            print(f"Error al cargar las imágenes: {e}")
            return None
        
        try:
            response = model.generate_content([prompt] + pil_images)
        
            if hasattr(response, 'text'):
                return response.text
            elif response and response.candidates and response.candidates[0] and hasattr(response.candidates[0], 'text'):
                return response.candidates[0].text
            else:
                print(f"Gemini no devolvió un texto válido para las imágenes.")
                return None
        except Exception as e:
            print(f"Error al procesar las imágenes {', '.join([os.path.basename(p) for p in image_paths])} con Gemini: {e}")
            return None
        finally:
            for img in pil_images:
                if img:
                    img.close()
                
    def export_to_docx(self):
        doc_dir = os.path.join(self.output_dir, "docs")
        os.makedirs(doc_dir, exist_ok=True)

        municipios = os.listdir(self.output_dir)
        for municipio in municipios:
            municipio_path = os.path.join(self.output_dir, municipio)
            if not os.path.isdir(municipio_path):
                continue

            doc = docx_document()
            # Titulo
            title = doc.add_heading(f"Informe de {municipio.capitalize()}", level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = title.add_run()
            run.bold = True
            run.underline = True
            run.font.name = 'Arial'
            run.font.size = Pt(30)



            for tab in sorted(os.listdir(municipio_path)):
                tab_path = os.path.join(municipio_path, tab)
                if not os.path.isdir(tab_path):
                    continue

                doc.add_heading(f"Tab {tab}", level=1)

                for file in sorted(os.listdir(tab_path)):
                    if file.endswith(".png"):
                        img_path = os.path.join(tab_path, file)

                        #doc.add_paragraph(os.path.splitext(file)[0])
                        doc.add_picture(img_path, width=Inches(6.0))
                        
                        try:
                            prompt = "Sos un analista de datos. Genera una profesional pero breve para el siguiente gráfico:"
                            description = MetabaseDashboardExtract.get_image_description_from_gemini([img_path], prompt)
                            time.sleep(0.5)
                            if description:
                                doc.add_paragraph(description.strip())
                            else:
                                doc.add_paragraph("(Fallo la descripción)")
                        except ResourceExhausted as e:
                            print(f"Se excedio la cuota de API")
                            raise
                        except Exception as e:
                            print(f"Error generando descripción para {file}: {e}")

            safe_municipio = municipio.replace(" ", "_").lower()
            output_path = os.path.join(doc_dir, f"{safe_municipio}.docx")
            doc.save(output_path)

    


    def docx_to_pdf(self):
        input_path = "output/docs/teulada.docx"
        output_dir = "output/pdf"

        # Asegurarse que la carpeta de salida exista
        os.makedirs(output_dir, exist_ok=True)

        try:
            subprocess.run([
                "D:/Programas/LibreOffice/program/soffice.exe",       
                "--headless",           # sin interfaz gráfica
                "--convert-to", "pdf",  # formato de salida
                "--outdir", output_dir, # carpeta destino
                input_path
            ], check=True)

            print("todo ok")
        except subprocess.CalledProcessError as e:
            print(f"error libreoffice: {e}")

    
    def run(self):
        inicio = time.time()
        try:
            self.login()
            #municipios = ["teulada", "benissa", "calpe", "javea", "mijas", "moraira", "platja_de_aro"]
            municipios = ["teulada"]
            #tabs = ["100-vuts-y-casas-rurales", "102-hoteles%2C-hostales-y-campings", "103-datos-de-fuentes-oficiales"]
            tabs = ["100-vuts-y-casas-rurales"]

            for municipio in municipios:
                for tab in tabs:
                    dashboard_url = f"https://analytics.peninsula.co/dashboard/32-{municipio}?a%25C3%25B1o=&fecha=&mes=&per%25C3%25ADodo=&poblaci%25C3%25B3n=teulada&poblaci%25C3%25B3n_%28consumo%29=calpe&poblaci%25C3%25B3n_2=benissa&tab={tab}&tipo_de_alojamiento=&tipo_de_establecimiento="
                    try:
                        self.capture_dashboard(dashboard_url, municipio, tab)
                    except Exception as e:
                        print(f"ERROR: {municipio} - {tab}: {e}")
                        break

            self.export_to_docx()
            self.docx_to_pdf()
            pdf_path = os.path.abspath("output/pdf/teulada.pdf")
            os.startfile(pdf_path)

        except ResourceExhausted as e:
            print(f"Se excedio la cuota de Gemini")
        except Exception as e:
            print(f"Error general: {e}")
        finally:
            self.driver.quit()
        fin = time.time()
        print(fin-inicio)

if __name__ == "__main__":
    email = os.getenv("METABASE_EMAIL")
    password = os.getenv("METABASE_PASSWORD")
    base_url = os.getenv("METABASE_BASE_URL")
    api_key = os.getenv("GEMINI_API_KEY")

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.0-flash')
    
    if not email or not password or not api_key or not base_url:
        raise ValueError("No estan las variables de entorno en el sistema")
        
    exporter = MetabaseDashboardExtract(
        email=email,
        password=password,
        base_url=base_url
    )
    exporter.run()