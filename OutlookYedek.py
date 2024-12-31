import os
import win32com.client
from datetime import datetime, timezone
import subprocess
import time
from datetime import datetime
# Kullanıcı adını dinamik olarak al
kullanici_adi = os.getlogin()
bilgisayar_adi = os.environ['COMPUTERNAME']
yedekleme_dosya_yolu = r"\\192.168.1.8\Ortak\ZAFER\Yedek\son_yedekleme.txt"
log = r"\\192.168.1.8\Ortak\ZAFER\Yedek\log.txt"

# Bilgisayar adını ve e-posta adresini bul
mail_adresi = None
son_yedekleme_zamani = None

# Geçerli satırı bul
dogru_satir = None
satirlar = []

with open(yedekleme_dosya_yolu, "r") as dosya:
    for satir in dosya:
        if bilgisayar_adi in satir and not satir.startswith("1"):  # "1" ile başlamayan satırı seç
            dogru_satir = satir.strip()  # Doğru olan satırı bul
            parcalar = dogru_satir.split(" | ")
            if len(parcalar) == 3:
                mail_adresi = parcalar[1]  # E-posta adresi
                son_yedekleme_zamani = datetime.strptime(parcalar[2], "%d.%m.%Y-%H:%M:%S").replace(tzinfo=timezone.utc)
        # Tüm satırları sakla
        satirlar.append(satir.strip())

# Ana ve yedek PST dosyalarının yollarını tanımla
ana_pst_dosya_yolu = f"C:\\Users\\{kullanici_adi}\\Documents\\Outlook Dosyaları\\{mail_adresi}.pst"
yedek_pst_dosya_yolu = rf"\\192.168.1.8\Ortak\ZAFER\Yedek\USERS\{bilgisayar_adi}\y{datetime.now().year} Yedek.pst"

# Outlook uygulamasını başlat
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# PST dosyasını ekle
try:
    namespace.AddStore(ana_pst_dosya_yolu)
except Exception as e:
    with open(log, "a") as log_dosyasi:
        hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Ana PST klasörü eklenirken hata: {e}\n")

# Yedek PST dosyasını ekle
try:
    namespace.AddStore(yedek_pst_dosya_yolu)

    # Yeni klasörleri kontrol et
    onceki_klasorler = [folder.Name for folder in namespace.Folders]  # Eski klasörler
    yedek_klasor = None  # Yedek klasör değişkenini başlat

    # Yedek klasörü kontrol et
    for folder in namespace.Folders:
        if folder.Name == f"y{datetime.now().year} Yedek":  # Yedek klasörün adı kontrol ediliyor
            yedek_klasor = folder
            break

    if not yedek_klasor:  # Eğer yedek klasör bulunamadıysa
        with open(log, "a") as log_dosyasi:
            hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
            log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Yedek PST klasörü bulunamadı.(Outlookda başında örn: y2024 Yedek formatında olması gerekiyor. \n")

except Exception as e:
    with open(log, "a") as log_dosyasi:
        hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Yedek PST klasörü eklenirken hata: {e}\n")

# Yedek klasör bulundu mu kontrol et
if 'yedek_klasor' in locals():  # yedek_klasor değişkeni tanımlandı mı
    try:
        ana_klasor = namespace.Folders[mail_adresi]

        # Kopyalama işlemi için fonksiyon
        def kopyala_klasorler(ana_klasor, yedek_klasor):
            # "Gelen Kutusu" altında alt klasörleri kontrol et
            if ana_klasor.Name == "Gelen Kutusu":
                hedef_klasor = yedek_klasor.Folders.Item(ana_klasor.Name) if ana_klasor.Name in [folder.Name for folder in yedek_klasor.Folders] else yedek_klasor.Folders.Add(ana_klasor.Name)
            else:
                hedef_klasor = yedek_klasor

            for alt_klasor in ana_klasor.Folders:
                # İlgili klasörleri hariç tut
                if alt_klasor.Name in ["Giden Kutusu", "Silinmiş Öğeler", "Taslaklar", "RSS Akışları", "Konuşma Eylemi Ayarları", "Hızlı Adım Ayarları", "Gereksiz E-posta","Önemsiz E-Posta"]:
                    continue

                # Alt klasördeki e-postaları kontrol et
                for mail in alt_klasor.Items:
                    try:
                        if mail.Class == 43:  # Sadece e-postalar
                            mail_zamani = mail.SentOn
                            if mail_zamani.tzinfo is None:
                                mail_zamani = mail_zamani.replace(tzinfo=timezone.utc)

                            if mail_zamani > son_yedekleme_zamani:
                                hedef_klasor_adi = alt_klasor.Name
                                if hedef_klasor_adi not in [folder.Name for folder in hedef_klasor.Folders]:
                                    hedef_alt_klasor = hedef_klasor.Folders.Add(hedef_klasor_adi)
                                else:
                                    hedef_alt_klasor = hedef_klasor.Folders.Item(hedef_klasor_adi)

                                # Hedef klasörde aynı mail var mı kontrol et
                                mail_var_mi = False
                                for hedef_mail in hedef_alt_klasor.Items:
                                    if (hedef_mail.Subject == mail.Subject and 
                                        hedef_mail.SentOn == mail.SentOn):
                                        mail_var_mi = True
                                        break

                                if not mail_var_mi:
                                    mail.Copy().Move(hedef_alt_klasor)

                    except Exception as e:
                        with open(log, "a") as log_dosyasi:
                            hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
                            log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Mail kopyalanırken hata: {mail.Subject}, hata: {e}\n")


                # Alt klasörleri kopyala
                try:
                    kopyala_klasorler(alt_klasor, hedef_klasor)
                except Exception as e:
                    with open(log, "a") as log_dosyasi:
                        hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
                        log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Alt klasör kopyalanırken hata: {alt_klasor.Name}, hata: {e}\n")

        # Kopyalama işlemini başlat
        kopyala_klasorler(ana_klasor, yedek_klasor)

        # Güncel tarih ve saat bilgisini yaz
        yeni_zaman = datetime.now().strftime('%d.%m.%Y-%H:%M:%S')
        if dogru_satir:
            for i, satir in enumerate(satirlar):
                if bilgisayar_adi in satir and not satir.startswith("1"):
                    satirlar[i] = f"{bilgisayar_adi} | {mail_adresi} | {yeni_zaman}"
                    break
            
            with open(yedekleme_dosya_yolu, "w") as dosya:
                for satir in satirlar:
                    dosya.write(satir + "\n")

    except Exception as e:
        with open(log, "a") as log_dosyasi:
            hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
            log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Kopyalama işlemi sırasında hata: {e}\n")

    finally:
        try:
            namespace.RemoveStore(yedek_klasor)
        except Exception as e:
            with open(log, "a") as log_dosyasi:
                hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
                log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Yedek PST dosyası kapatılırken hata: {e}\n")

else:
    with open(log, "a") as log_dosyasi:
        hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Yedek klasör bulunamadığı için kopyalama işlemi gerçekleştirilemiyor.\n")

try:
    subprocess.run(["schtasks", "/run", "/tn", "OuSearchService"], check=True)
except subprocess.CalledProcessError as e:
    with open(log, "a") as log_dosyasi:
        hata_zamani = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        log_dosyasi.write(f"[{hata_zamani}] - {mail_adresi} - Görev çalıştırılırken hata: {e}\n")  