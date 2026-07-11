import os
import pandas as pd

file_name = "rekap_data.xlsx"

columns = [
    "No",
    "Nama Siswa",
    "Nama Ayah",
    "Nama Ibu",
    "Alamat",
    "Jenis Kelamin",
    "Diterima di Kelas",
    "Tempat Lahir",
    "Tanggal Lahir",
    "Keluar dari Tk/Paud",
    "Status Murid",
    "Keterangan"
]

datalist = []

print("==== Rekap Data Siswa ====")
print("ketik 'exit' untuk keluar dari program.\n")

while True:
    print(f"-- Silahkan masukan data siswa ke-{len(datalist) + 1} ---")

    nama_siswa = input("Nama siswa: ")
    if nama_siswa.lower() == "exit":
        break
    ayah = input("Masukan Nama ayah: ")
    ibu = input("Masukan Nama ibu: ")
    alamat = input("Masukan alamat: ")
    jk = input("Jenis Kelamin: ")
    kelas = input("Diterima di kelas: ")
    tempat = input("Tempat lahir: ")
    tgl_lahir = input("Tanggal lahir (DD-MM-YYY): ")
    asal_tk = input("Keluar dari TK/PAUD: ")
    status = input("Status (diterima/tidak): ")
    keterangan = input("Keterangan: ")

    data_siswa = {
    "No": len(datalist) + 1,
    "Nama Siswa": nama_siswa,
    "Nama Ayah": ayah,
    "Nama Ibu": ibu,
    "Alamat": alamat,
    "Jenis Kelamin": jk,
    "Diterima di Kelas": kelas,
    "Tempat Lahir": tempat,
    "Tanggal Lahir": tgl_lahir,
    "Keluar dari Tk/Paud": asal_tk,
    "Status Murid": status,
    "Keterangan": keterangan
}

    datalist.append(data_siswa)
    print("-> Data berhasil disimpan sementara.\n")
    #tanya user mau lanjut atau udahan
    lanjut = input("Mau input data siswa lagi? (y/n): ")
    if lanjut.lower() != "y":
        break

#Proses penyimpanan data ke Excel
if datalist:
    new_df = pd.DataFrame(datalist)

    if os.path.exists(file_name):
        old_df = pd.read_excel(file_name, engine="openpyxl")

        #hitung ulang nomor urut data
        start_no = len(old_df) + 1 
        for i, data in enumerate(datalist):
            new_df.at[i, "No"] = start_no + i

            # Gabungkan data lama dengan data baru
        final_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        # Kalau file belum ada, pakai data baru ini aja
        final_df = new_df

    # Atur urutan kolom di Excel biar pas sesuai list columns
    final_df = final_df.reindex(columns=columns)

    # Ekspor ke file Excel tanpa membawa index bawaan pandas
    final_df.to_excel(file_name, index=False, engine="openpyxl")
    print(f"\n[SUKSES] Berhasil menyimpan {len(datalist)} data baru ke '{file_name}'!")
else:
    print("\nTidak ada data yang disimpan.")

