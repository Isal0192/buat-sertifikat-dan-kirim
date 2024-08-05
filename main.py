import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
from docx2pdf import convert


def get_dataframe(path):
    """
    Baca file excel yang berisi data pendaftaran seminar, lalu ambil kolom yang dipilih user.
    Kemudian tampilkan data yang memiliki nilai kolom tersebut.

    """

    #membaca file excel
    df = pd.read_excel(path)
    array = df.columns.to_list()

    #menampilkan nama column untuk di ambil columnnya
    print("pilih salah satu")
    for i in range(len(array)):
        print(f"[{i}]. {array[i]}")
    
    #mengambil kolom yang dipilih user
    choice = int(input("\npilihan: "))
    selected_column = array[choice]
    print(f"kolom yang dipilih: {selected_column}")
    list_coich = df[selected_column]

    print(f"\nJumlah data yang memiliki nilai {selected_column}: {len(list_coich)}\ntelah di tambahkan ke list")
    print()
    return list_coich

def read_sertifikat(path, array):
    """
    membaca dox dan menulis dalam dari array yang di pilih

    """

    sertif = DocxTemplate(path)
    
    for i in array:
        sertif.render({"NAMA": i})
        sertif.nama = f"sertifikat {i}.docx"
        sertif.save(f"sertifikat/{sertif.nama}")

    convert('sertifikat')
    return "sertifikat berhasil di buat..."

import os
def send_email(path):
    """
    Kirim email ke peserta.
    """

    #mencari sertifikat pdf di folder
    files = [f for f in os.listdir("sertifikat") if f.endswith(".pdf")]
    get_email = get_dataframe(path).to_string(index=False)
    get_email.lstrip()
    email = get_email.split("\n")
    # email.reverse()
    

    
    dictionary = [[files[i],email[i]] for i in range(len(files))]
    print(dictionary)



def main():

    #mangambil data dari file excel 
    print("*"*40)
    path = input("masukan path excel: \n")
    template_sertif = input("masukan path templat sertif: \n")
    df = get_dataframe(path)
    print("*"*40)

    #membuat duplikat data
    with open("data.txt","w") as data:
        data.write(df.to_string(index=False))
        print("data yang di pilih berhasil disalin dalam file data.txt...\n")
    
    #membuat sertifikat berdasarkan data yang ada
    array = [i for i in df]
    print("*"*40)
    print(read_sertifikat(template_sertif,array))

    #memngirim ke email
    # print("*"*40,"\n\nmengirim ke email perserta....\n pilih email")
    # send_email(path)

if __name__ == "__main__":
    main()
    # send_email("excel/pendaftaran_seminar.xlsx")
    