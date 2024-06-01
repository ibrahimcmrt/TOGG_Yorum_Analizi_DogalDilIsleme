from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
import time

# YouTube video URL'si
video_url = "https://www.youtube.com/watch?v=D0OBxfjw5ec"

# Chrome WebDriver'ı başlat
chrome_options = Options()
chrome_options.add_argument("--headless")  # Arka planda çalışması için
driver = webdriver.Chrome(options=chrome_options)

# Video sayfasını aç
driver.get(video_url)
time.sleep(5)  # Sayfanın tamamen yüklenmesini bekle

# Yorumları tutacak bir liste oluştur
yorumlar = []

# Sayfanın sonuna kadar inerek yorumları topla
while len(yorumlar) < 500:
    # Yorumları bul
    yorum_elementleri = driver.find_elements(By.CSS_SELECTOR, "#content-text")
    # Yorumları ekle
    for element in yorum_elementleri:
        yorumlar.append(element.text)
        if len(yorumlar) >= 500:
            break
    # Sayfayı aşağı kaydır
    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
    time.sleep(3)  # Yeni yorumların yüklenmesini bekle

# WebDriver'ı kapat
driver.quit()

# Excel dosyasını yükle
dosya_yolu = "/Users/ibrahimcomert/Desktop/YoutubeYorum.xlsx"  # Excel dosyasının yolu
workbook = load_workbook(dosya_yolu)
sheet = workbook.active

# Yorumları Excel dosyasına yaz
for i, yorum in enumerate(yorumlar, start=2201):
    sheet[f'B{i}'] = yorum

# Değişiklikleri kaydet
workbook.save(dosya_yolu)
