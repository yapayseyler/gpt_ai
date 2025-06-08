# Codex Calculator ve Excel Arayüzü

Bu proje, basit bir hesap makinesi modülünün yanı sıra Excel dosyalarındaki TC sorguları için örnek bir Tkinter arayüzü içerir.

## Özellikler

- Toplama, çıkarma, çarpma ve bölme fonksiyonları (calculator.py)
- Excel dosyasından TC kimlik bilgisi arayarak sonuçları ikinci sayfaya yazan basit bir arayüz (`excel_gui.py`)

## calculator.py Kullanımı

```python
import calculator

print(calculator.add(2, 3))  # 5
print(calculator.divide(10, 2))  # 5.0
```

## excel_gui.py Kullanımı

1. `Excel Yükle` düğmesi ile iki sayfadan oluşan Excel dosyasını seçin. İlk sayfada A sütunu TC, B sütunu Ad Soyad, C sütunu Kan Grubu olmalıdır.
2. TC numarası girip `Sorgula` butonuna bastığınızda, ilgili bilgiler bulunur ve hem arayüzde gösterilir hem de varsa ikinci sayfanın A2, B2 ve C2 hücrelerine yazılır.

Arayüzün çalışması için `openpyxl` paketinin kurulu olması gerekir.
