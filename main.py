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