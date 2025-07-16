import os
import time
import zipfile
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from docx import Document
from docx.shared import Inches

# Iniciar sesion en Metabase
# Capturar imagenes de los graficos que estan en el dashboard
# Guardar las imagenes con el titulo del grafico como nombre de archivo
# Comprimir la carpeta en un ZIP

class MetabaseDashboardExporter:
    def __init__(self, email, password, base_url, dashboard_url, output_dir="screenshots"):
        self.email = email
        self.password = password
        self.base_url = base_url
        self.dashboard_url = dashboard_url
        self.output_dir = output_dir
        

        # Configuración del navegador Chrome en modo headless
        # Comportamiento de Chrome antes de abrirlo
        chrome_options = Options()
        # Sin interfaz grafica
        chrome_options.add_argument("--headless=new")
        # Aceleracion por hardware de la GPU (Deshabilitado)
        chrome_options.add_argument("--disable-gpu")
        # Como esta en headless, le damos una resolucion para que renderice 
        chrome_options.add_argument("--window-size=1920,1080")
        # Deshabilitamos el directorio del sistema (si usamos Docker nos da problemas), vamos con disco
        #chrome_options.add_argument("--disable-dev-shm-usage")
        # En Docker puede fallar el sanbox
        #chrome_options.add_argument("--no-sandbox")

        # Desactivamos notificaciones 
        prefs = {
            "profile.default_content_setting_values.notifications": 2, # 1 (notis), 2, (bloqueadas), 3 (preguntar)
            "profile.default_content_settings.popups": 0 
        }
        chrome_options.add_experimental_option("prefs", prefs)

        # Navegador
        self.driver = webdriver.Chrome(options=chrome_options)
        # Carpeta para las imagenes
        os.makedirs(self.output_dir, exist_ok=True)

    # Login pero ademas hay que navegar al dashboard 
    def login(self):
        
        login_url = f"{self.base_url}/auth/login"
        self.driver.get(login_url)

        try:
            # Espera y completa el formulario de login
            # Esperamos hasta que una condicion se cumpla, o 20 segundos

            username_field = WebDriverWait(self.driver, 20).until(
                # Expected Conditions (EC)
                                            # Por atributo
                EC.element_to_be_clickable((By.NAME, "username"))
            )

            password_field = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "password"))
            )

            # Escribimos en los campos
            username_field.send_keys(self.email)
            password_field.send_keys(self.password)


            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
            )
            login_button.click()

            # Espera a que termine el login (cambia la URL)
            WebDriverWait(self.driver, 30).until(
                EC.url_contains("/")
            )

            # Va al dashboard directamente
            self.driver.get(self.dashboard_url)

            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="dashcard-container"]'))
            )

        except TimeoutException as e:
            print(f"Error en login: {e}")
            self.driver.save_screenshot("login_error.png")
            raise

    def capture_dashboard(self):
        """Captura los gráficos del dashboard como imágenes individuales"""
        time.sleep(10)  # Espera a que renderice todo

        # Encuentra los contenedores de gráficos
        cards = self.driver.find_elements(By.CSS_SELECTOR, '[data-testid="dashcard-container"]')

        if not cards:
            print("No se encontraron gráficos.")
            self.driver.save_screenshot("dashboard_debug.png")
            return

        for i, card in enumerate(cards):
            try:
                # Hace scroll al gráfico para asegurarse que está visible
                # scrollIntoView() funcion de JS que desplaza la pagina hasta que el elemento sea visible
                # {block: 'center'} alinea el centro del elemento con el centro del viewport (verticalmente)- start/end/center
                # arguments[0] es para pasar el card desde python a JS
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card)
                time.sleep(2)

                # Espera visibilidad
                WebDriverWait(self.driver, 10).until(EC.visibility_of(card))

                title_element = card.find_element(By.CSS_SELECTOR, "[data-testid='legend-caption-title']")
                title = title_element.text.strip()

                # Limpieza de caracteres inválidos para nombre de archivo
                # Expresiones regulares (regex) es una forma avanzada de buscar y reemplazar texto. 
                # modulo re
                # re.sub(patron, reemplazo, texto)
                # en el patron r'' es un raw string, para que python no interprete los caracteres especiales

                clean_title = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", title)[:80]
                filename = f"{clean_title}.png"
                # screenshots/filename
                filepath = os.path.join(self.output_dir, filename)

                card.screenshot(filepath)
                print(f"Capturado: {filepath}")

            except Exception as e:
                print(f"Error capturando gráfico {i+1}: {e}")
                continue
    # Crea un archivo ZIP con todas las imagenes capturadas
    def zip_screenshots(self, zip_name="dashboard_screenshots.zip"):
        
        try:

            with zipfile.ZipFile(zip_name, "w") as zipf:

                for filename in os.listdir(self.output_dir):

                    if filename.endswith(".png"):

                        filepath = os.path.join(self.output_dir, filename)

                        zipf.write(filepath, arcname=filename)
            print(f"ZIP generado: {zip_name}")

        except Exception as e:
            print(f"Error creando ZIP: {e}")

    def export_to_word(self, doc_name="dashboard.docx"):
        doc = Document()
        # self.output_dir = screenshots
        for filename in sorted(os.listdir(self.output_dir)):
            if filename.endswith(".png"):

                img_path = os.path.join(self.output_dir, filename)
                doc.add_paragraph(filename.replace(".png", ""))
                doc.add_picture(img_path, width=Inches(8))
                # Aca se puede poner texto que haya sido analizado con la API de Gemini
                #doc.add_page_break()
        doc.save(doc_name)
        print(f"Word generado: {doc_name}")

    def run(self):
        #Ejecuta todo el proceso completo de login, captura y compresion
        try:
            self.login()
            self.capture_dashboard()
            self.zip_screenshots()
            self.export_to_word()
        except Exception as e:
            print(f"Error general: {e}")
        finally:
            self.driver.quit()



if __name__ == "__main__":
    email = os.getenv("METABASE_EMAIL")
    password = os.getenv("METABASE_PASSWORD")
    base_url = os.getenv("METABASE_BASE_URL")
    dashboard_url = os.getenv("METABASE_DASHBOARD_URL")

    if not email or not password:
        raise ValueError("No estan las variables de entorno en el sistema")
    exporter = MetabaseDashboardExporter(
        email=email,
        password=password,
        dashboard_url=dashboard_url,
        base_url=base_url
    )
    exporter.run()
