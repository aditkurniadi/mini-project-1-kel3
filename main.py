# MINI PROJECT 1 - KELOMPOK 3
# BERIKAN KOMENTAR PADA SETIAP BAGIAN CODE YANG DI TAMBAHKAN

# Bagian Programer / Data Analyst
import pandas as pd

# import dan baca data Excel
file_path = "Data Wisudawan Kel 3.xlsx"
data = pd.read_excel(file_path)

# Hitung jumlah wisudawan per prodi
jumlah_per_prodi = data.groupby('Program Studi')['NIM'].count().reset_index()
jumlah_per_prodi.rename(columns={'NIM': 'Jumlah Wisudawan'}, inplace=True)

print("=== Jumlah Wisudawan per Program Studi ===")
print(jumlah_per_prodi)
print()

# Klasifikasi Grade IPK
def klasifikasi_grade(ipk):
    if ipk >= 3.75:
        return 'A'
    elif ipk >= 3.50:
        return 'B+'
    elif ipk >= 3.00:
        return 'B'
    elif ipk >= 2.50:
        return 'C'
    else:
        return 'D'

data['Grade'] = data['IPK'].apply(klasifikasi_grade)

# Klasifikasi Predikat Wisuda
def klasifikasi_predikat(row):
    ipk = row['IPK']
    lama = row['Lama Studi (Semester)']

    if ipk >= 3.75 and lama <= 8:
        return 'Cumlaude (Dengan Pujian)'
    elif ipk >= 3.50 and lama <= 9:
        return 'Sangat Memuaskan'
    elif ipk >= 3.00:
        return 'Memuaskan'
    else:
        return 'Cukup'

data['Predikat Wisuda'] = data.apply(klasifikasi_predikat, axis=1)

# menampilkan data hasil klasifikasi
print("=== Data Wisudawan dengan Grade dan Predikat ===")
print(data[['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Grade', 'Predikat Wisuda']].head(10))

# Simpan hasil analisis ke file Excel baru
output_file = "hasil_analisis_wisudawan.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name='Data Lengkap', index=False)
    jumlah_per_prodi.to_excel(writer, sheet_name='Jumlah Per Prodi', index=False)

print(f"\nFile hasil analisis disimpan sebagai: {output_file}")


# BAGIAN VISUAL DESIGNER / DATA VISUALIZATION
# ⿡ GRAFIK JUMLAH WISUDAWAN PER PRODI
plt.figure(figsize=(8, 5))                                
bars = plt.bar(jumlah_per_prodi['Program Studi'],          
               jumlah_per_prodi['Jumlah Wisudawan'],      
               color="#0BEADE")                            

plt.title('Jumlah Wisudawan per Program Studi', fontsize=14, fontweight='bold')
plt.xlabel('Program Studi')
plt.ylabel('Jumlah Wisudawan')
plt.xticks(rotation=25)                                    

# Menampilkan nilai jumlah di atas setiap batang
for bar in bars:
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
             str(bar.get_height()), ha='center', va='bottom', fontsize=10)

plt.tight_layout()                                         #
plt.show()                                               


# ⿢ GRAFIK DISTRIBUSI PREDIKAT WISUDA (PIE CHART)
predikat_counts = data['Predikat Wisuda'].value_counts()   
colors = ["#E42092", "#FC672D", "#FD0808", "#F5FD08"]     

plt.figure(figsize=(6, 6))
plt.pie(predikat_counts,
        labels=predikat_counts.index,                      
        autopct='%1.1f%%',                            
        startangle=90,                                     
        colors=colors,
        textprops={'fontsize': 10})                        

plt.title('Distribusi Predikat Wisuda', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.show()


# ===========================================================
# Bagian Penyimpanan Hasil Analisis ke Excel
# ===========================================================
output_file = "hasil_analisis_wisudawan.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name='Data Lengkap', index=False)
    jumlah_per_prodi.to_excel(writer, sheet_name='Jumlah Per Prodi', index=False)
    predikat_counts.to_frame('Jumlah').to_excel(writer, sheet_name='Distribusi Predikat', index=False)

print(f"\nFile hasil analisis disimpan sebagai: {output_file}")


#=====================SELESAI=========================

