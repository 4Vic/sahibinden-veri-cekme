import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import os

# kütüphaneleri kontrol eder, yoksa yükler
os.system("pip install selenium")
os.system("pip install openpyxl")

print("""       



         d8b 888    888               888                                                 d88P  d8888  888     888 d8b          
         Y8P 888    888               888                                                d88P  d8P888  888     888 Y8P          
             888    888               888                                               d88P  d8P 888  888     888              
 .d88b.  888 888888 88888b.  888  888 88888b.       .d8888b  .d88b.  88888b.d88b.      d88P  d8P  888  Y88b   d88P 888  .d8888b 
d88P"88b 888 888    888 "88b 888  888 888 "88b     d88P"    d88""88b 888 "888 "88b    d88P  d88   888   Y88b d88P  888 d88P"    
888  888 888 888    888  888 888  888 888  888     888      888  888 888  888  888   d88P   8888888888   Y88o88P   888 888      
Y88b 888 888 Y88b.  888  888 Y88b 888 888 d88P d8b Y88b.    Y88..88P 888  888  888  d88P          888     Y888P    888 Y88b.    
 "Y88888 888  "Y888 888  888  "Y88888 88888P"  Y8P  "Y8888P  "Y88P"  888  888  888 d88P           888      Y8P     888  "Y8888P 
     888                                                                                                                        
Y8b d88P                                                                                                                        
 "Y88P"                                                                                                                         

                                                                                                  
""")
time.sleep(3)

sayfa_numarasi = 0 # Her Sayfada 50 ürün olduğu için bu sayı 50 ve katları olarak yükselecek

workbook = Workbook() # Excel dosyası oluştur

kolon = workbook.active # Sütun başlıklarını yaz

kolon["A1"] = "Marka"
kolon["B1"] = "Seri"
kolon["C1"] = "Model"
kolon["D1"] = "İlan Başlığı"
kolon["E1"] = "Yıl"
kolon["F1"] = "KM"
kolon["G1"] = "Renk"
kolon["H1"] = "Fiyat"
kolon["I1"] = "İlan Tarihi"
kolon["J1"] = "İl/İlçe"

driver = webdriver.Chrome() # Web tarayıcıyı aç

driver.get("https://www.sahibinden.com/otomobil?pagingSize=50")

# Bilgileri çek

while True:
    for i in range(1, 51):
        try:
            if i == 50:
                sayfa_numarasi += 50

            # Marka bilgisini al
            marka_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[2]")
            marka = marka_bilgisi.text
        except:
            marka = "Veri Yok"

        try:
            # Model bilgisini al
            model_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[4]")
            model = model_bilgisi.text
        except:
            model = "Veri Yok"

        try:
            # İlan Başlığını al
            ilan_basligi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[5]/a")

            ilan = ilan_basligi.text
        except:
            ilan = "Veri Yok"

        try:
            # Yıl bilgisini al
            yil_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[6]")
            yil = yil_bilgisi.text
        except:
            yil = "Veri Yok"

        try:

            #il/ilçe bilgisini al
            il_ilce_bilgisi_al = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[11]")
            il_ilce_bilgisi_yazi = il_ilce_bilgisi_al.text

        except:

            #il_ve_ilce = "Veri Yok"
            ilce_bilgisi_yazi = "Veri Yok"

        try:
            # Renk  bilgisini al
            renk_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[8]")
            renk = renk_bilgisi.text
        except:
            renk = "Veri Yok"

        try:
            # İlan tarihi bilgisini al

            # günü al
            ilan_tarihi_bilgisi_gun = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[10]/span[1]")

            ilan_tarihi_gun = ilan_tarihi_bilgisi_gun.text

            # yılı al
            ilan_tarihi_bilgisi_yil = driver.find_element(By.XPATH,f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[10]/span[2]")

            ilan_tarihi_yil = ilan_tarihi_bilgisi_yil.text

            ilan_tarihi_gun_yil = ilan_tarihi_gun+" "+ilan_tarihi_yil

        except:

            ilan_tarihi_gun_yil = "Veri Yok"

        try:
            # Seri bilgisini al
            seri_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[3]")
            seri = seri_bilgisi.text
        except:
            seri = "Veri Yok"

        try:
            # Kilometre bilgisini al
            km_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[7]")
            km = km_bilgisi.text
        except:
            km = "Veri Yok"

        try:
            # Fiyat bilgisini al
            fiyat_bilgisi = driver.find_element(By.XPATH, f"/html/body/div[5]/div[4]/form/div/div[3]/table/tbody/tr[{i}]/td[9]/div")
            fiyat = fiyat_bilgisi.text
        except:
            fiyat = "Veri Yok"


        print(f"\nMarka: {marka} \nSeri:{seri} \nModel:{model} \nİlan Başlığı:{ilan} \nYıl:{yil} \nKilometre:{km} \nRenk:{renk}   \nFiyat:{fiyat}   \nİlan Tarihi:{ilan_tarihi_gun_yil} \nİl/İlçe:{il_ilce_bilgisi_yazi} \n{'-' * 50}")


        # Verileri excel dosyasına ekle

        kolon.append([marka, seri, model, ilan, yil, km, renk, fiyat, ilan_tarihi_gun_yil, il_ilce_bilgisi_yazi])

    print("Sonraki sayfaya geçiş yapılıyor...")

    try:
        print("Sayfa Numarası :", sayfa_numarasi)
        driver.get(f"https://www.sahibinden.com/otomobil?pagingOffset={sayfa_numarasi}&pagingSize=50")     # Sonraki sayfaya geçiş yap

        time.sleep(5) # Sayfanın yüklenmesi için 5 saniye bekle

    except:

        print("Son sayfaya ulaşıldı.\nExcel dosyasına kaydedildi.")
        workbook.save(filename="sahibinden-arac-bilgileri.xlsx") # verileri excel dosyasına kaydet
        break