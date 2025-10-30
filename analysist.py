# ===============================
#  Analisis Penjualan Produk
# Oleh: Muhammad Hafirst Firdaus
# ===============================

import pandas as pd
import matplotlib.pyplot as plt
import os

# 1️ Baca file
file_path = "Project_Analisis_Penjualan_Pertama.xlsx"
data = pd.read_excel(file_path)

# 2️ Informasi dasar
print("=== 5 Data Pertama ===")
print(data.head(), "\n")

print("=== Informasi Data ===")
print(data.info(), "\n")

print("=== Jumlah Data Kosong ===")
print(data.isnull().sum(), "\n")

# 3️ Statistik deskriptif
print("=== Statistik Deskriptif ===")
print(data.describe(), "\n")

# 4️ Total dan rata-rata penjualan
if 'Total' in data.columns:
    total_semua = data['Total'].sum()
    rata2 = data['Total'].mean()
    print(f"💰 Total seluruh penjualan: Rp{total_semua:,.0f}")
    print(f"📈 Rata-rata penjualan: Rp{rata2:,.0f}\n")

# 5️ Pastikan folder untuk hasil visualisasi ada
os.makedirs("hasil_visualisasi", exist_ok=True)

# 6️ Produk terlaris
if 'Jenis Produk' in data.columns and 'Total' in data.columns:
    produk_terlaris = data.groupby('Jenis Produk')['Total'].sum().sort_values(ascending=False)
    print("=== Produk Terlaris ===")
    print(produk_terlaris.head(10), "\n")

    plt.figure(figsize=(10,5))
    produk_terlaris.head(10).plot(kind='bar', color='#4C8BF5')
    plt.title("Top 10 Produk dengan Penjualan Tertinggi", fontsize=13)
    plt.xlabel("Jenis Produk")
    plt.ylabel("Total Penjualan (Rp)")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.savefig("hasil_visualisasi/top10_produk.png")
    plt.show()
    plt.close()

# 7️ Tren penjualan per bulan
if 'Tanggal' in data.columns and 'Total' in data.columns:
    data['Tanggal'] = pd.to_datetime(data['Tanggal'])
    data['Bulan'] = data['Tanggal'].dt.to_period('M')
    tren_bulan = data.groupby('Bulan')['Total'].sum()

    print("=== Tren Penjualan per Bulan ===")
    print(tren_bulan, "\n")

    plt.figure(figsize=(10,5))
    plt.plot(tren_bulan.index.astype(str), tren_bulan.values,
             marker='o', color='#2AA876', linewidth=2)
    plt.title("Tren Penjualan Bulanan", fontsize=13)
    plt.xlabel("Bulan")
    plt.ylabel("Total Penjualan (Rp)")
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig("hasil_visualisasi/tren_bulanan.png")
    plt.show()
    plt.close()

print(" Analisis selesai! Grafik disimpan di folder 'hasil_visualisasi/'.")
