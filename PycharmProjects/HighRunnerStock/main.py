import pandas as pd
import os

def print_progress(percentage):
    print(f"\rİlerleme: %{percentage * 100}", end="")

# Dinamik dosya yolu
dosya_yolu = os.path.join(os.path.expanduser("~"), "Desktop", "zsd0500.xlsx")

print_progress(0.1)

# Excel dosyasını pandas ile oku
df = pd.read_excel(dosya_yolu, engine='openpyxl')

print_progress(0.3)

# Ağırlığı ton cinsine dönüştürme
df["Net ağırlık (ton)"] = df["Net ağırlık"] / 1000

print_progress(0.5)

# Malzeme ve tarih bazında unique kayıtları bulma
unique_shipments = df.drop_duplicates(subset=["Malzeme", "Fiili mal hrkt.trh."])

print_progress(0.6)

# Frekans bazında sıralama
frequency_df = unique_shipments["Malzeme"].value_counts().reset_index()
frequency_df.columns = ["Malzeme", "Sevk Frekansı"]

print_progress(0.7)

# Ağırlık (ton cinsinden) bilgisini ekleyerek sonucu oluşturma ve müşteri bilgisini ekleyerek tamamlama
weight_df = df.groupby("Malzeme").agg({'Net ağırlık (ton)': 'sum'}).reset_index()
most_common_customer = df.groupby("Malzeme")["Siparişi Veren Adı"].apply(lambda x: x.value_counts().index[0]).reset_index()
most_common_customer.columns = ["Malzeme", "En Çok Sevk Edilen Müşteri"]

print_progress(0.9)

merged_df = frequency_df.merge(weight_df, on="Malzeme").merge(most_common_customer, on="Malzeme")
merged_df = merged_df.sort_values(by="Sevk Frekansı", ascending=False)

print("\nSevk Frekansı Bazında High Runner Ürünler:")
print(merged_df.head(10)[["Malzeme", "Sevk Frekansı", "Net ağırlık (ton)", "En Çok Sevk Edilen Müşteri"]])  # İlk 10 ürünü gösterir

print_progress(1.0)
