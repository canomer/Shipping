import docx
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

# .docx dosyasını aç
doc = docx.Document("belge.docx")

# Tablo sayacı
table_count = 0

# Her tabloyu işle
for table in doc.tables:
    table_count += 1
    total_rows = len(table.rows)
    
    # Her satırı işle
    for row in table.rows:
        # İlk 3 hücreyi koru, son hücreyi temizle
        for i in range(3, len(row.cells)):
            for paragraph in row.cells[i].paragraphs:
                paragraph.clear()
        
        # Hücre içeriğini düzenle, eğer 3 hücre boş değilse, % hesapla ve ek olarak birleştirme işlemi yap
        if row.cells[0].text.strip() and row.cells[1].text.strip() and row.cells[2].text.strip():
            percentage = 100
            row.cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row.cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row.cells[2].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        else:
            percentage = 0
        
        # Yüzdeyi son hücreye ekle
        row.cells[-1].text = f"{percentage}%"

# Dosyayı kaydet
doc.save("duzenlenmis_belge.docx")

print(f"Toplam {table_count} tablo düzenlendi ve dosya kaydedildi.")
