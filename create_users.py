import requests, sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook

# ===== Настройки ============================================================
API_BASE = "http://localhost:5678"
API_KEY = "your_api_key_here" # Замените на ваш API ключ
USERS_NUM = 5
USER_PREFIX = "student_"
PASSWORD = "Password1"
EMAIL_DOM = "training.local"

HEADERS = {
    "X-N8N-API-KEY": API_KEY,
    "Content-Type": "application/json",
}
USERS_URL = f"{API_BASE.rstrip('/')}/api/v1/users"


def setup_selenium():
    options = Options()
    options.add_argument("--headless=new")  # Новый режим headless
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    # Убедитесь, что ChromeDriver установлен правильно
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def accept_invitation_selenium(invite_url: str, login: str, password: str) -> bool:
    driver = setup_selenium()
    try:
        driver.get(invite_url)
        print("Page title:", driver.title)  # Для отладки
        
        # Ожидаем загрузки всей страницы
        WebDriverWait(driver, 15).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # Альтернативные селекторы для полей формы
        first_name = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='firstName'], #firstName, [data-test-id='firstName'] input"))
        )
        
        last_name = driver.find_element(By.CSS_SELECTOR, "input[name='lastName'], #lastName, [data-test-id='lastName'] input")
        passwd = driver.find_element(By.CSS_SELECTOR, "input[name='password'], #password, [data-test-id='password'] input")
        
        # Заполняем форму
        first_name.clear()
        first_name.send_keys(login)
        
        last_name.clear()
        last_name.send_keys(login)  # Пустая фамилия
        
        passwd.clear()
        passwd.send_keys(password)
        
        # Скриншот перед отправкой (для отладки)
        driver.save_screenshot(f"before_submit_{login}.png")
        
        # Нажимаем кнопку
        submit_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-test-id='form-submit-button'], .submit-btn"))
        )
        submit_btn.click()
        
        # Ждём изменения URL или появления элемента подтверждения
        try:
            WebDriverWait(driver, 15).until(
                lambda d: "workflow" in d.current_url.lower() or "login" in d.current_url.lower()
            )
            print(f"User {login} successfully activated")
            return True
        except:
            print("Possible success - saving final state screenshot")
            driver.save_screenshot(f"after_submit_{login}.png")
            return True
            
    except Exception as e:
        print(f"[SELENIUM ERROR] {str(e)}")
        driver.save_screenshot(f"error_{login}.png")
        return False
    finally:
        driver.quit()

def create_user(login: str, password: str) -> str:
    user_body = [{
        "email": f"{login}@{EMAIL_DOM}",
        "firstName": login,
        "lastName": "",
        "password": password,
        "globalRoleId": 2,
    }]
    
    r = requests.post(USERS_URL, headers=HEADERS, json=user_body, timeout=15)
    
    if r.ok:
        print(f"User {login} created successfully")
        invite_url = r.json()[0]["user"]["inviteAcceptUrl"]
        
        # Попытка активации через браузер
        if accept_invitation_selenium(invite_url, login, password):
            return invite_url
        return invite_url  # Возвращаем ссылку даже если автоматизация не сработала
    else:
        print(f"[ERROR] {login}: {r.status_code} {r.text}", file=sys.stderr)
        return ""

def main() -> None:
    # Создаем новую Excel-книгу
    wb = Workbook()
    ws = wb.active
    
    # Устанавливаем заголовки
    ws.append(["Email", "Password"])
    
    for i in range(1, USERS_NUM + 1):
        username = f"{USER_PREFIX}{i}"
        email = f"{username}@{EMAIL_DOM}"
        
        # Создаем пользователя
        invite_link = create_user(username, PASSWORD)
        
        # Если пользователь успешно создан, добавляем в Excel
        if invite_link:
            ws.append([email, PASSWORD])
            print(f"Added user: {email} | {PASSWORD}")
    
    # Сохраняем файл
    wb.save("users.xlsx")
    print("Data successfully saved to users.xlsx")

if __name__ == "__main__":
    main()
    