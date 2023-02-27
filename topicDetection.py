"""
Versi 2.0 - January 2023

@author: Gerald Plakasa
"""

import xlwings as xw
import pandas as pd
import numpy as np
import yake
import string
import nltk
import re
import os
from nltk.corpus import stopwords
from nltk import word_tokenize
from nltk.stem import PorterStemmer
from deep_translator import GoogleTranslator
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import CountVectorizer
import en_core_web_sm
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from datetime import datetime
import time

pd.options.mode.chained_assignment = None

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('maxent_ne_chunker')
nltk.download('words')
nltk.download('stopwords')

# menyimpan kata stopword
stopword = stopwords.words('english')

# membuat object untuk Stemming
ps = PorterStemmer()

# ambil data yang ingin di proses
def getData(path, nama_sheet, sheet, nama_kolom):

    try:
        wb1 = xw.Book(path)
    except:
        sheet["A1"].color = "#f54257"
        raise Exception("Path tidak ditemukan")

    sheet["A1"].color = "#6ff542"

    try:
        wbs = wb1.sheets[nama_sheet]
    except:
        sheet["A2"].color = "#f54257"
        raise Exception("Sheet tidak ditemukan")

    sheet["A2"].color = "#6ff542"

    try:
        df = wbs.range((1, 1)).options(pd.DataFrame, header=1, index=False, expand='table', empty=np.NaN).value
        df[[nama_kolom]]
    except:
        sheet["A3"].color = "#f54257"
        raise Exception("Nama Kolom Tidak ditemukan")

    sheet["A3"].color = "#6ff542"

    return wbs

# fungsi untuk membenarkan typo
def fix_typo(df_teks, nama_kolom, kata_typo):
    # setting list kosong untuk menyimpan hasil
    hasil = []
    # Ulangi untuk setiap teks di kolom "Reason for Score"
    for teks in df_teks[nama_kolom]:
        # Jika teks bukan kosong atau NaN
        if not pd.isna(teks):
            teks = str(teks)
            try:

                # Tambahkan spasi setelah titik
                teks = re.sub(r'(?<=[.])(?=[^\\s])', r' ', teks)
                # hapus spasi jika ada sebelum titik
                teks = teks.replace(" .", ".")
                
                # Jadikan 3 atau lebih karakter berulang menjadi 2
                pattern = re.compile(r"(.)\1{1,}", re.DOTALL)
                teks = pattern.sub(r"\1\1", teks)
            except:
                raise Exception(teks)

            # lakukan tokenisasi dan masukkan ke teks_temp
            teks_temp = word_tokenize(str(teks))
            # ulangi untuk setiap token pada teks_temp
            for token in tuple(teks_temp):
                # Cek jika token merupakan kata yang ada pada daftar kunci kata typo maka replace dengan yang benar
                if token.lower() in kata_typo.keys():
                    # ambil kata benar / value di dictionary
                    value = kata_typo[token.lower()]
                    # Replace token dengan value
                    teks = teks.replace(token, value)
            # lalu hasilnya masukkan ke list hasil
            hasil.append(teks)
        else:
            # Jika teks kosong atau NaN masukkan list hasil dengan np.NaN
            hasil.append(np.NaN)

    return hasil

# membuat sebuah fungsi untuk text to file
def createFile(df_teks, n_awal, n_akhir, nama_kolom):
    # mengambil data yang tidak kosong dari index n_awal sampai n_akhir
    # lalu di simpan dalam bentuk list pada variabel hasil
    hasil = (df_teks.loc[df_teks[nama_kolom].notnull()]).loc[n_awal:n_akhir, nama_kolom].to_list()
    kata_file = ""

    # membuat file baru bernama translate.txt
    with open('translate.txt', 'w', encoding="utf-8") as f:
        # menggabungkan seluruh list teks menjadi 1 teks di kata_file
        kata_file += '\n'.join(str(x) for x in hasil)
        # menyimpan seluruh list teks ke file yang sudah dibuat
        f.write(kata_file)

    # mengembalikan nilai panjang teks di kata_file
    return len(kata_file)

# fungsi untuk melakukan translate
def translate(df_teks, n, nama_kolom):
    n_awal = 0
    n_akhir = 0
    hasil = []

    # Jika n kurang dari 200 maka panjang batasnya sebesar n,
    # jika tidak panjang batasnya 200
    if n < 200:
        panjang_batas_n = n
    else:
        panjang_batas_n = 200

    # Lakukan Perulangan jika index awal n masih kurang dari n
    while(n_awal < n):

        # Setting index akhir merupakan index akhir + panjang batas n
        n_akhir += panjang_batas_n

        # mengirimkan nilai index awal, index akhir, dan dataframe ke fungsi createFile
        # hasilnya akan mendapatkan panjang kata dari index awal sampai index akhir
        panjang_karakter = createFile(df_teks.copy(), n_awal, n_akhir, nama_kolom)

        pengurang_panjang = 0
        # melakukan perulangan lagi jika panjang kata pada 200 baris masih >= 5000
        while panjang_karakter >= 5000:

            # jika index awal masih kurang dari index akhir - 50
            # maka coba kurangi index akhirnya 50
            # dan tambahkan pengurang panjangnya 50
            if n_awal < n_akhir - 50:
                pengurang_panjang += 50
                n_akhir = n_akhir - 50
            # jika index awal masih kurang dari index akhir - 25
            # maka coba kurangi index akhirnya 25
            # dan tambahkan pengurang panjangnya 25
            elif n_awal < n_akhir - 25:
                pengurang_panjang += 25
                n_akhir = n_akhir - 25
            # jika index awal masih kurang dari index akhir - 5
            # maka coba kurangi index akhirnya 5
            # dan tambahkan pengurang panjangnya 5
            elif n_awal < n_akhir - 5:
                pengurang_panjang += 5
                n_akhir = n_akhir - 5
            # jika tidak memenuhi semua kurangi index akhir dengan 1
            # dan tambahkan pengurang panjangnya 1
            else:
                pengurang_panjang += 1
                n_akhir = n_akhir - 1

            # coba lagi data dengan jumlah row yang sudah di ubah dijadikan file
            # akan mendapatkan panjang kata yang baru
            panjang_karakter = createFile(df_teks.copy(), n_awal, n_akhir, nama_kolom)

        # Jika sudah mendapatkan panjang kata < 5000, baru lakukan translate pada file translate.txt
        hasil_translate = GoogleTranslator(source='id', target='en').translate_file('translate.txt')
        # hasil teks dipisah menjadi list
        list_translate = hasil_translate.split('\n')

        count = 0
        # selanjutnya proses memasukkan teks yang sudah di tranlate ke list hasil
        for teks in df_teks[nama_kolom][n_awal:n_akhir]:
            if teks == teks:
                hasil.append(list_translate[count])
                count += 1
            else:
                hasil.append(np.NaN)

        # index awal merupakan index awal sekarang ditambahkan dengan panjang batas yang sudah dikurangi dengan pengurang panjang
        n_awal = n_awal + panjang_batas_n - pengurang_panjang

    try:
        os.remove("translate.txt")
    except:
        pass

    return hasil

# melakukan lower pada hasil translate
def translateLower(df_teks):

    hasil_temp = df_teks['Translate2'].astype(str).str.lower()
    hasil = hasil_temp.replace('nan', np.NaN)

    return hasil

# proses ekstraksi kata kunci
def keywordExtraction(df_teks):

    # set jumlah maksimal suku kata
    max_ngram_size = 2
    # set Threshold penghapusan Kata kunci Duplikat
    deduplication_threshold = 0.9
    # membuat object yake dengan parameter yang diseting
    kw_extractor = yake.KeywordExtractor(lan="en", n=max_ngram_size, dedupLim=deduplication_threshold, top=20, features=None)

    # Buat untuk menyimpan hasil kata kunci
    keywords_w = []
    keywords = []

    for teks in df_teks['Translate']:
        # Jika teks tidak kosong ambil kata kuncinya
        if teks == teks:
            # Ekstrak kata kunci dari teks
            keyword_w = kw_extractor.extract_keywords(str(teks))
            keyword = []
            # lakukan perulangan untuk setiap kata kunci
            for kw in keyword_w:
                # pisahkan nama kata kuncinya dan bobotnya
                nama, bobot = zip(kw)
                # lakukan proses penghapusan karakter punctuation pada teks
                nama_punctuation = ''.join([word for word in nama[0] if word not in string.punctuation])
                # lakukan proses stemming pada kata kuncinya
                nama_stem = ps.stem(nama_punctuation)
                # simpan hasil yang sudah di stemming ke keyword
                keyword.append(nama_stem)

            # jika hasil ekstraksi kata kunci kosong
            # simpan dengan nama teksnya dan bobot 0
            if len(keyword_w) < 1:
                keyword_w = [(teks, 0)]
                keyword = [teks]
        else:
            keyword_w = teks
            keyword = teks
        # simpan untuk seluruh kata kunci pada teks di variabel keywords dan keywords_w
        keywords_w.append(keyword_w)
        keywords.append(keyword)

    return (keywords, keywords_w)

# membuat fungsi untuk memberihkan sebuah kata atau string
def clean_string(text):
    # Pembersihan karakter spesial
    text = ''.join([word for word in text if word not in string.punctuation])
    # membuat teks menjadi lower
    text = text.lower()
    # penghapusan kata stopword
    text = ' '.join([word for word in word_tokenize(text) if word not in stopword])
    # penghapusan spasi berlebih di awal dan akhir kalimat
    text = text.strip()

    return text

# melakukan kategorisasi
def categorize(df_teks, df_kategori, n):
    # Setting variabel yang dibutuhkan
    hasil_kategori = []
    count = 0
    batas_awal_n = 0
    batas_akhir_n = 0
    n_batas = 2000

    # Ulangi sampai batas awal n lebih dari n
    while batas_awal_n <= n:

        # setiap perulangan batas akhir n di tambah 2000
        batas_akhir_n += n_batas

        # Mengambil semua kata kunci pada 2000 data
        kata_kunci_all = list(df_teks['Keywords'][batas_awal_n:batas_akhir_n].dropna().explode())

        # menyimpan panjang kata kunci all sebelum di tambah
        panjang_kata_kunci_all = len(kata_kunci_all)

        # menambahkan isi kata kunci all dengan kata kunci yang ada pada kamus kata kunci
        kata_kunci_all += list(df_kategori['Keywords'])

        # membersihkan seluruh kata kunci pada fungsi yang didefinisikan sebelumnya
        cleaned = list(map(clean_string, kata_kunci_all))
        # mengubah kata kunci yang telah bersih menjadi vektor
        vectorizer = CountVectorizer().fit_transform(cleaned)
        # menjadikan vectorizernya array
        vectors = vectorizer.toarray()
        # melihat similarity semua kata kunci kepada masing-masing kata kunci
        csim = cosine_similarity(vectors)
        
        count = 0
        # melakukan perulangan untuk setiap kumpulan kata kunci pada 2000 data
        for kata_kunci_row in df_teks['Keywords'][batas_awal_n:batas_akhir_n]:
            
            # Jika kata kunci tidak kosong lakukan proses
            if kata_kunci_row == kata_kunci_row:
                # ambil panjang ada berapa kata kunci pada row tersebut
                panjang = len(kata_kunci_row)
                # simpan count_next dengan count saat ini + banyaknya kata kunci
                count_next = count + panjang
                # Seting sebuah set dalam variabel kategori
                kategori = set()
                kategori_count = []
                # Perulangan pada panjang count sampai count next
                for i in range(count, count_next):
                    # ambil index dimana nilai similaritynya lebih dari 0,7
                    idx_temp = np.where(csim[i]>0.7)
    
                    # ulangi setiap index temp yang di set sebelumnya
                    # Jika index temp nilainya lebih dari panjang kata kunci sebelumnya, masukkan ke index
                    idx = [j for j in idx_temp[0] if j >= panjang_kata_kunci_all]
                
                    # ulangi untuk setiap index yang sudah di set sebelumnya
                    for k in idx:
                        # ambil letak index sebelumnya berdasarkan indexnya di dataframe kamus kata kunci
                        index_kategori = df_kategori.loc[df_kategori['Keywords'] == cleaned[k]].index
                        
                        # jika index pada kamus kata kunci ada, masukkan kata kuncinya ke set kategori
                        if len(index_kategori) > 0:
                            kategori_count.append(df_kategori['Category'][index_kategori[0]])
                            kategori.add(df_kategori['Category'][index_kategori[0]])
                
                # Jika kategorinya hanya 1, kategorinya hanya Coverage, terdapat kata slow pada kata kuncinya, maka proses
                if (len(kategori) == 1) and ("Coverage" in kategori) and ("slow" in kata_kunci_row):
                    # lakukan perulangan untuk setiap kata kunci pada row tersebut
                    for kata in kata_kunci_row:
                        # jika contains kata slow, maka tambahkan kategori Data
                        if "slow" in kata:
                            kategori.add("Data")
                            kategori_count.append("Data")

                # jika kategori kosong, maka kategorinya "Non Categorize"
                if len(kategori) == 0:
                    kata_sendiri = ("slow", "slower", "connect", "disconnect", "slowli", "antislow", "reconnect")
                    cek_non_categorize = True
                    # Cek terlebih dahulu untuk yang indikasi non categorize
                    # apakah terdapat kata sendiri di kata kuncinya
                    # jika iya tambahkan kategori data
                    for kata in kata_sendiri:
                        if kata in kata_kunci_row:
                            kategori.add("Data")
                            kategori_count.append("Data")
                            cek_non_categorize = False
                    # jika kategori tetap non categorize maka tambahkan kategori non categorize
                    if cek_non_categorize:
                        kategori.add("Non Categorize")
                        kategori_count.append("Non Categorize")
                
                # lakukan perhitungan untuk berapa jumlah kemunculan kategori
                kategori_dict = {i:kategori_count.count(i) for i in kategori_count}
                
                # untuk location check selalu jadikan 1
                if "Location Check" in tuple(kategori_dict.keys()):
                    kategori_dict["Location Check"] = 1
                        
                # hasil akhir kategori di simpan pada hasil_categori
                hasil_kategori.append(kategori_dict)
                # set nilai countnya menjadi nilai count next
                count = count_next
            else:
                hasil_kategori.append(kata_kunci_row)
        
        # tambah nilai batas awal n dengan 2000
        batas_awal_n += n_batas
        
    return hasil_kategori

# membuat fungsi untuk melakukan stemming pada seluruh teks di list
def stemProcess(list_kata):
    # untuk setiap kata di list lakukan proses steming dan strip
    hasil = [ps.stem(kata).strip() for kata in list_kata]
    
    # kembalikan list yang telah di lakukan proses
    return hasil

# melakukan proses stability
def setStability(df_teks, df_stability, n):
    hasil_stability = []
    # ulangi untuk setiap teks yang sudah di translate
    for i, teks in enumerate(df_teks['Translate']):
        # jika teks tidak kosong lakukan proses
        if teks == teks:
            # seting variabel stabil dengan kosong/NaN terlebih dahulu
            stabil = np.NaN
            # Jika termasuk kategori Call, Gaming, Data, atau Coverage maka proses
            if "Gaming" in df_teks['Kategori'][i] or "Data" in df_teks['Kategori'][i] or "Coverage" in df_teks['Kategori'][i] or "Call" in df_teks['Kategori'][i]:
                # jika kata di kamus stabil data pada hasil tokenisasi kata
                # maka variabel stabil dijadikan "Stability"
                if len(set(df_stability["Keywords"]) & set(stemProcess(word_tokenize(teks.lower())))) != 0:
                    stabil = "Stability"
            # hasil akhirnya masukkan pada list hasil Stability
            hasil_stability.append(stabil)
        else:
            hasil_stability.append(np.NaN)
    
    return hasil_stability

# membuat fungsi untuk melakukan pengurutan 2 list berdasarkan 1 list
def sort_list(list1, list2):

    maks = max(list2)

    indicates = [i for i, x in enumerate(list2) if x == maks]
    zipped_pairs = zip(list2, list1)
    z = [x for _, x in sorted(zipped_pairs)]
    return z, indicates

# melakukan proses finalisasi kategori
def finalKategori(df_teks, kepentingan_category):
    # buat variabel list kosong untuk menyimpan hasil kategori dan location check
    final_hasil = []
    location_check = []
    # ulangi setiap kumpulan kategori yang ada pada kolom kategori
    for i, categorys in enumerate(df_teks['Kategori']):
        # jika kumpulan kategori tidak kosong lakukan proses
        if categorys == categorys:
            # jika isi kumpulan kategori lebih dari satu lakukan proses
            if len(categorys) > 1:

                # ambil jumlah dari masing-masing kategori
                jumlah = list(categorys.values())

                # melakukan sorting 2 list kategori dan jumlahnya
                # dengan pengurutan berdasarkan jumlah kategori
                categorys_hasil, indicates = sort_list(list(categorys.keys()), jumlah)

                # ambil kategori dengan jumlah terbanyak
                category_final = categorys_hasil[len(categorys_hasil)-1]
                
                # Jika ada jumlah yang sama 
                if len(indicates) > 1:
                    # ambil nama-nama kategori dengan jumlah sama
                    categorys_temp = [(list(categorys.keys()))[idx] for idx in indicates]
                    # ambil tingkat untuk masing-masing nama kategori
                    tingkat_temp = [kepentingan_category[category] for category in categorys_temp]
                    # urutkan kategori berdasarkan tingkat kepentingannya
                    categorys_hasil, indicates = sort_list(categorys_temp, tingkat_temp)
                    # ambil nilai dengan kepentingan tertinggi
                    category_final = categorys_hasil[0]
                    
                    # Jika ada Non_Network=(Coverage, Data) jadikan Network_Product
                    if (category_final == "Pricing" or category_final == "Reward" or category_final == "Product") and ("Coverage" in categorys_hasil or "Data" in categorys_hasil):
                        category_final = "Network_Product"

                # ambil nilai dengan kepentingan tertinggi
                # jika location Check terdapat dalam kategori lakukan proses
                if "Location Check" in categorys_hasil and (category_final != "Pricing" and category_final != "Reward" and category_final != "Product"):
                    # Jika kategori tertinggi bukan location check lakukan proses
                    if category_final != "Location Check":
                        category_final = category_final + ", " + "Location Check"
                    # Jika terdapat location Check, jadikan true
                    location_check.append(True)
                else:
                    location_check.append(False)
                
            # jika isi kumpulan kategori hanya satu ambil kategorinya untuk dijadikan final kategori
            else:
                category_final = list(categorys)[0]
                # Jika Location Check terdapat dalam kategori, jadikan true
                if "Location Check" in categorys:
                    location_check.append(True)
                else:
                    location_check.append(False)
            
            # Jika Kategori final merupakan Call dan terdapat kata CS di sentimentnnya
            # maka jadikan kategorinya Non_Network
            if (category_final == "Call") and ("CS" in word_tokenize(df_teks["Translate"][i])):
                category_final == "Product"
            
            # Jika pada baris tersebut merupakan stability maka tambahkan stability pada final kategori
            if df_teks['Stability'][i] == df_teks['Stability'][i]:
                if "Call" in word_tokenize(category_final) or "Gaming" in word_tokenize(category_final) or "Data" in word_tokenize(category_final) or "Coverage" in word_tokenize(category_final):
                    category_final = category_final + ", " + "Stability"      
        else:
            category_final = categorys
            location_check.append(False)
        
        final_hasil.append(category_final)
    
    return (final_hasil, location_check)

def getTopic(kategori_list):
    hasil_topic = []
    for kategori in kategori_list:
        if kategori == kategori:
            kategori_temp = kategori.split(",")
            if "Call" in kategori_temp:
                hasil_topic.append("Call")
            elif "Network_Product" in kategori_temp:
                hasil_topic.append("Network_Product")
            elif "Pricing" in kategori_temp:
                hasil_topic.append("Pricing")
            elif "Reward" in kategori_temp:
                hasil_topic.append("Reward")
            elif "Product" in kategori_temp:
                hasil_topic.append("Product")
            elif "Gaming" in kategori_temp:
                hasil_topic.append("Game")
            elif "Data" in kategori_temp:
                hasil_topic.append("Data")
            elif "Coverage" in kategori_temp:
                hasil_topic.append("Coverage")
            elif "Non Categorize" in kategori_temp:
                hasil_topic.append("Non Categorize")
            else:
                hasil_topic.append("Location Check")
        else:
            hasil_topic.append("empty")
    
    return hasil_topic

# melakukan ekstraksi lokasi
def extractLocation(df_teks):
    nlp = en_core_web_sm.load()

    location = []
    
    # Ulangi untuk setiap row pada kolom Location Check
    for i, check_location in enumerate(df_teks['Location Check']):
        # Jika Row valuenya True lakukan proses
        if check_location:
            Location_Entity = []
            kurung = ["(", ")", "[", "]", "{", "}"]
            
            # hapus setiap teks yang mengandung karakter kurung di row yang sama
            teks = ''.join([word for word in df_teks["Translate2"][i] if word not in kurung])
            # hapus setiap teks yang mengandung kata stopword
            teks = ' '.join([word for word in word_tokenize(teks) if word not in stopword])
            # Terapkan en_core_web_sm pada teks
            doc = nlp(teks)
            
            # ulangi untuk setiap entitas yang ter ekstrak dari teks
            for entity in doc.ents:
                # Set kata-kata yang bukan lokasi
                bukan_lokasi = ["xl", "axis", "promotions", "pln", "signal", "sony", "4g+", "app", 'axisnet', "youtube", "mayxl"
                               'sandiaga uno', 'satisfied', 'rp', 'tai', 'telkomsel', 'id', 'hbs', "xl signal", "gb", "bad", "subhanallah",
                               "quota", 'myxl', "xl network", 'cs', 'fast beach', 'strengthen', 'indosat', '100k']
                # jika label dari entitas terbetu GPE, LOC, PERSON, ORG, NORP, FAC lakukan proses
                if entity.label_ == 'GPE' or entity.label_ == 'LOC' or entity.label_ == 'PERSON' or entity.label_ == 'ORG' or entity.label_ == 'NORP' or entity.label_ == 'FAC':
                    # Jika teks dari entity tidak ada dalam daftar bukan lokasi, masukan ke dalam list
                    if not entity.text.lower() in bukan_lokasi:
                        Location_Entity.append(entity.text)
            # jika setelah di ulang lokasi tidak ditemukan, maka masukan value kosong
            if len(Location_Entity) < 1:
                Location_Entity = np.NaN
        else:
            Location_Entity = np.NaN 
        # simpan seluruh entitas lokasi yang di ekstrak pada list location setiap rownya
        location.append(Location_Entity)
    
    return location

def sentiment_scores(sentence):

    sid_obj = SentimentIntensityAnalyzer()
    sentiment_dict = sid_obj.polarity_scores(sentence)

    if sentiment_dict['compound'] >= 0.05 :
        return "Positive"
    elif sentiment_dict['compound'] <= - 0.05 :
        return "Negative"
    else :
        return "Neutral"

def setSentiment(df_teks):
    return [sentiment_scores(teks) if teks == teks else np.NaN for teks in df_teks['Translate']]

# main proses
def main():

    start = time.time()
    
    # Ambil datetime Sekarang
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    
    # ambil excel vba utama
    wb = xw.Book.caller()
    sheet = wb.sheets["Sheet1"]
    history = wb.sheets["history1"]
    
    # Jadikan Background Biru
    sheet["A1"].color = "#6674f2"
    sheet["A2"].color = "#6674f2"
    sheet["A3"].color = "#6674f2"
    sheet["A4"].color = "#6674f2"
    
    # hapus bagian loading proses
    sheet["A6"].color = None
    sheet["A6"].value = None
    sheet["A7"].value = None
    
    # ambil data path, sheet, dan nama kolom input user
    path = sheet["A1"].value
    nama_sheet = sheet["A2"].value
    nama_kolom = sheet["A3"].value
    do_sentiment = sheet["A4"].value
    
    with_sentiment = "Tanpa Sentiment Analysis"
    
    # Ambil data
    wb1 = getData(path, nama_sheet, sheet, nama_kolom)
    
    kolom_terakhir = wb1["A1"].expand("right").last_cell.column
    
    rows = []
    for i in range(1,kolom_terakhir):
        last_row = wb1.range((1, i)).expand("table").last_cell.row
        rows.append(last_row)
    lokasi_kolom = rows.index(max(rows))
            
    # ubah data menjadi pandas dataframe
    df = wb1.range((1, lokasi_kolom+1)).options(pd.DataFrame, header=1, index=False, expand='table', empty=np.NaN).value
    df = df.replace(r'^\s*$', np.nan, regex=True)
    df = df.fillna(value=np.nan)
    
    # Ambil data sentimen
    df_non_nan = df.copy()
    df_non_nan = df_non_nan.dropna(subset=nama_kolom)
    
    # Ambil Panjang Awal
    panjang_awal = df.shape[0]
    
    n = (df.shape)[0]
    df_teks = df[[nama_kolom]]
    
    sheet["A6"].value = "Process..."
    sheet["A6"].font.color = "#000000"
    sheet["A6"].color = "#f08035"
    
    # ambil data Typo di sheet typo
    sheet1 = wb.sheets["typo"]
    # jadikan pandas dataframe
    df_typo = sheet1.range('A1').options(pd.DataFrame,header=1,index=False, expand='table').value
    # dataframe to dictionary
    kata_typo = df_typo.set_index('typo').to_dict()['benar']
    
    # fix typo
    simpan = fix_typo(df_teks.copy(), nama_kolom, kata_typo)
    df_teks.loc[:,nama_kolom] = simpan
    
    # Translate Data
    simpan = translate(df_teks.copy(), n, nama_kolom)
    df_teks['Translate2'] = simpan
    
    simpan = translateLower(df_teks.copy())
    df_teks['Translate'] = simpan
    
    # Keyword Extraction
    keywords, keywords_w = keywordExtraction(df_teks.copy())

    # pengubahan data keyword dari list ke string
    simpan = []
    for kw in keywords:
        if kw == kw:
            simpan_temp = ", ".join(str(x) for x in kw)
            simpan_temp = "[ "+simpan_temp+" ]"
        else:
            simpan_temp = np.NaN
        simpan.append(simpan_temp)
        
    # masukan keyword ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Keywords"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(simpan)
    df_teks['Keywords'] = keywords
    
    # tambah tulisan 25%
    sheet["A6"].value = "Process 25%..."
    sheet["A6"].font.color = "#000000"
    sheet["A6"].color = "#f08035"
    
    # ambil data kategori di sheet category
    sheet1 = wb.sheets["category"]
    # jadikan pandas dataframe
    df_kategori = sheet1.range('A1').options(pd.DataFrame,header=1,index=False, expand='table').value
    # Lakukan Stemming pada seluruh kata kunci
    df_kategori['Keywords'] = df_kategori['Keywords'].apply(lambda x: ps.stem(x))
    
    # proses kategorisasi
    kategori = categorize(df_teks.copy(), df_kategori.copy(), n)
    
    # pengubahan data kategori dari list ke string
    simpan = [str(ktgr) if ktgr == ktgr else np.NaN for ktgr in kategori]
    
    # masukan kategori ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Kategori"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(simpan)
    df_teks['Kategori'] = kategori
    
    # tambah tulisan 50%
    sheet["A6"].value = "Process 50%..."
    sheet["A6"].font.color = "#000000"
    sheet["A6"].color = "#f08035"
    
    # ambil data stability di sheet stability
    sheet2 = wb.sheets["stability"]
    # jadikan pandas dataframe
    df_stability = sheet2.range('A1').options(pd.DataFrame,header=1,index=False, expand='table').value
    # Lakukan Stemming pada seluruh kata kunci
    df_stability['Keywords'] = df_stability['Keywords'].apply(lambda x: ps.stem(x))

    # proses set stability
    stability = setStability(df_teks.copy(), df_stability.copy(), n)
    
    # masukan stability ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Stability"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(stability)
    df_teks['Stability'] = stability
    
    # tambah tulisan 75%
    sheet["A6"].value = "Process 75%..."
    sheet["A6"].font.color = "#000000"
    sheet["A6"].color = "#f08035"
    
    # ambil data Typo di sheet typo
    sheet1 = wb.sheets["category_level"]
    # jadikan pandas dataframe
    df_tingkat_kepentingan = sheet1.range('A1').options(pd.DataFrame,header=1,index=False, expand='table').value
    # dataframe to dictionary
    kepentingan_category = df_tingkat_kepentingan.set_index('Category').to_dict()['Level']
    
    # proses finalisasi kategori dan ambil location check
    final_hasil, location_check = finalKategori(df_teks.copy(), kepentingan_category)
    
    df_teks['Final Kategori'] = final_hasil
    
    df_teks['Location Check'] = location_check
    
    # proses ambil Topics NLP
    topic_nlp = getTopic(final_hasil)
    
    df_teks['Topics_NLP'] = topic_nlp
    
    # proses ekstraksi lokasi
    location = extractLocation(df_teks.copy())
    
    # pengubahan lokasi dari list ke string
    simpan = []
    for kw in location:
        if kw == kw:
            simpan_temp = ", ".join(str(x) for x in kw)
            simpan_temp = "[ "+simpan_temp+" ]"
        else:
            simpan_temp = np.NaN
        simpan.append(simpan_temp)
    
    # simpan lokasi ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Location"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(simpan)
    
    df_teks['Location'] = simpan
    
    if do_sentiment.lower() == 'y':
        
        with_sentiment = "Dengan Sentiment Analysis"
        
        # set sentiment
        sentiment = setSentiment(df_teks.copy())
        # masukan sentiment Analysis ke excel data
        last_cell = wb1["A1"].expand("right").last_cell.column
        wb1.range((1,last_cell+1)).options(index=False).value = "Sentiment Analysis"
        wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(sentiment)
    
    # masukan final kategori ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Final Kategori (NLP)"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(final_hasil)
    
    # masukan topics_NLP ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Topics_NLP"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(topic_nlp)
    
    last_row = history.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
    last_value = history.range((last_row,1)).options(index=False).value
    if last_value == "No":
        no = 1
    else:
        no = int(last_value)+1
    
    end = time.time()
    
    path_list = path.split("\\")
    
    history.range((last_row+1,1)).options(index=False, header=False).value = pd.DataFrame([[no, 
                                                                                            dt_string,
                                                                                            path_list[len(path_list)-1],
                                                                                            "Topic Detection", 
                                                                                            str(panjang_awal) + " Data", 
                                                                                            str(len(df_non_nan)) + " Data", 
                                                                                            str(round((end-start)/60, 2)) + " Menit", 
                                                                                            with_sentiment]])
    
    sheet["A6"].value = "Done!"
    sheet["A6"].font.color = "#000000"
    sheet["A6"].color = "#84de1d"
    
    sheet["A7"].value = str(round((end-start)/60, 2)) + " Menit"
    
def bersihkan():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.clear_contents()
    sheet["A1"].value = r"C:\Users\gwx1201968\Documents\Sentimen Analysis\xlwings test\xwfix1\data.xlsx"
    sheet["A2"].value = "Sheet1"
    sheet["A3"].value = "Reason for Score"
    sheet["A4"].value = "n"
    sheet["A11"].value = r"C:\Users\gwx1201968\Documents\Sentimen Analysis\xlwings test\xwfix1\data.xlsx"
    sheet["A12"].value = "Sheet1"
    sheet["A13"].value = "Reason for Score"
    sheet["D11"].value = r"C:\Users\gwx1201968\Documents\Sentimen Analysis\xlwings test\xwfix1\data.xlsx"
    sheet["D12"].value = "Sheet1"
    sheet["D13"].value = "Month, MSISDN, Score, Reason for Score"
    sheet["A1"].color = "#6674f2"
    sheet["A2"].color = "#6674f2"
    sheet["A3"].color = "#6674f2"
    sheet["A4"].color = "#6674f2"
    sheet["A11"].color = "#6674f2"
    sheet["A12"].color = "#6674f2"
    sheet["A13"].color = "#6674f2"
    sheet["D11"].color = "#6674f2"
    sheet["D12"].color = "#6674f2"
    sheet["D13"].color = "#6674f2"
    sheet["B1"].value = "<<< Masukan Path lokasi data"
    sheet["B2"].value = "<<< Nama Sheet"
    sheet["B3"].value = "<<< Nama Kolom Sentiment"
    sheet["B4"].value = "<<< Masukkan Proses Sentimen Analysis (y/n)"
    sheet["B11"].value = "<<< Masukan Path lokasi data"
    sheet["B12"].value = "<<< Nama Sheet"
    sheet["B13"].value = "<<< Nama Kolom Sentiment"
    sheet["E11"].value = "<<< Masukan Path lokasi data"
    sheet["E12"].value = "<<< Nama Sheet"
    sheet["E13"].value = "<<< Nama Kolom Patokan (Pisahkan dengan koma)"
    sheet["D1"].value = r"contoh = C:\Users\gwx1201968\Documents\Sentimen Analysis\xlwings test\xwfix1\data.xlsx"
    sheet["D2"].value = "contoh = Sheet1"
    sheet["D3"].value = "contoh = Reason for Score"
    sheet["A6"].color = None
    sheet["A15"].color = None
    sheet["D15"].color = None

# ambil data yang ingin di proses
def getData2(path, nama_sheet, sheet, nama_kolom):
    try:
        wb1 = xw.Book(path)
    except:
        sheet["A11"].color = "#f54257"
        raise Exception("Path tidak ditemukan")

    sheet["A11"].color = "#6ff542"
    
    try:
        wbs = wb1.sheets[nama_sheet]
    except:
        sheet["A12"].color = "#f54257"
        raise Exception("Sheet tidak ditemukan")
    
    sheet["A12"].color = "#6ff542" 
    
    try:
        df = wbs.range((1, 1)).options(pd.DataFrame, header=1, index=False, expand='table', empty=np.NaN).value
        df[[nama_kolom]]
    except:
        sheet["A13"].color = "#f54257"
        raise Exception("Nama Kolom Tidak ditemukan")
    
    sheet["A13"].color = "#6ff542" 
        
    return wbs

def SentimentAnalysis():
    
    start = time.time()
    
    # ambil excel vba utama
    wb = xw.Book.caller()
    sheet = wb.sheets["Sheet1"]
    history = wb.sheets["history1"]
    
    # Ambil datetime Sekarang
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    
    # Jadikan Background Biru
    sheet["A11"].color = "#6674f2"
    sheet["A12"].color = "#6674f2"
    sheet["A13"].color = "#6674f2"
    
    # hapus bagian loading proses
    sheet["A15"].color = None
    sheet["A15"].value = None
    sheet["A16"].value = None
    
    # ambil data path, sheet, dan nama kolom input user
    path = sheet["A11"].value
    nama_sheet = sheet["A12"].value
    nama_kolom = sheet["A13"].value
    
    # Ambil data
    wb1 = getData2(path, nama_sheet, sheet, nama_kolom)
    
    kolom_terakhir = wb1["A1"].expand("right").last_cell.column
    
    rows = []
    for i in range(1,kolom_terakhir):
        last_row = wb1.range((1, i)).expand("table").last_cell.row
        rows.append(last_row)
    lokasi_kolom = rows.index(max(rows))
            
    # ubah data menjadi pandas dataframe
    df = wb1.range((1, lokasi_kolom+1)).options(pd.DataFrame, header=1, index=False, expand='table', empty=np.NaN).value
    df = df.replace(r'^\s*$', np.nan, regex=True)
    df = df.fillna(value=np.nan)
    
    # Ambil data sentimen
    df_non_nan = df.copy()
    df_non_nan = df_non_nan.dropna(subset=nama_kolom)
    
    # Ambil Panjang Awal
    panjang_awal = df.shape[0]
    
    sheet["A15"].value = "Process..."
    sheet["A15"].font.color = "#000000"
    sheet["A15"].color = "#f08035"
    
    n = (df.shape)[0]
    df_teks = df[[nama_kolom]]
    
    # ambil data Typo di sheet typo
    sheet1 = wb.sheets["typo"]
    # jadikan pandas dataframe
    df_typo = sheet1.range('A1').options(pd.DataFrame,header=1,index=False, expand='table').value
    # dataframe to dictionary
    kata_typo = df_typo.set_index('typo').to_dict()['benar']
    
    # fix typo
    simpan = fix_typo(df_teks.copy(), nama_kolom, kata_typo)
    df_teks.loc[:,nama_kolom] = simpan
    
    # Translate Data
    simpan = translate(df_teks.copy(), n, nama_kolom)
    df_teks['Translate2'] = simpan
    
    simpan = translateLower(df_teks.copy())
    df_teks['Translate'] = simpan
    
    # tambah tulisan 50%
    sheet["A15"].value = "Process 50%..."
    sheet["A15"].font.color = "#000000"
    sheet["A15"].color = "#f08035"
    
    # set sentiment
    sentiment = setSentiment(df_teks.copy())
    # masukan sentiment Analysis ke excel data
    last_cell = wb1["A1"].expand("right").last_cell.column
    wb1.range((1,last_cell+1)).options(index=False).value = "Sentiment Analysis"
    wb1.range((2,last_cell+1)).options(index=False).value = pd.Series(sentiment)
    
    last_row = history.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
    last_value = history.range((last_row,1)).options(index=False).value
    if last_value == "No":
        no = 1
    else:
        no = int(last_value)+1
    
    end = time.time()
    
    path_list = path.split("\\")
    
    history.range((last_row+1,1)).options(index=False, header=False).value = pd.DataFrame([[no, 
                                                                                            dt_string,
                                                                                            path_list[len(path_list)-1],
                                                                                            "Sentimen Analysis", 
                                                                                            str(panjang_awal) + " Data", 
                                                                                            str(len(df_non_nan)) + " Data", 
                                                                                            str(round((end-start)/60, 2)) + " Menit", 
                                                                                            "-"]])
    
    sheet["A15"].value = "Done!"
    sheet["A15"].font.color = "#000000"
    sheet["A15"].color = "#84de1d"
    
    sheet["A16"].value = str(round((end-start)/60, 2)) + " Menit"

# ambil data yang ingin di proses
def getData3(path, nama_sheet, sheet, list_patokan):
    try:
        wb1 = xw.Book(path)
    except:
        sheet["D11"].color = "#f54257"
        raise Exception("Path tidak ditemukan")

    sheet["D11"].color = "#6ff542"
    
    try:
        wbs = wb1.sheets[nama_sheet]
    except:
        sheet["D12"].color = "#f54257"
        raise Exception("Sheet tidak ditemukan")
    
    sheet["D12"].color = "#6ff542" 
    
    try:
        df = wbs.range((1, 1)).options(pd.DataFrame, header=1, index=False, expand='table', empty=np.NaN).value
        df[list_patokan]
    except:
        sheet["D13"].color = "#f54257"
        raise Exception("Nama Kolom Tidak ditemukan")
    
    sheet["D13"].color = "#6ff542" 
        
    return wbs
    
def RemoveDuplicated():
    
    start = time.time()
    
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    
    # ambil excel vba utama
    wb = xw.Book.caller()
    sheet = wb.sheets["Sheet1"]
    history = wb.sheets["history1"]
    
    # Jadikan Background Biru
    sheet["D11"].color = "#6674f2"
    sheet["D12"].color = "#6674f2"
    sheet["D13"].color = "#6674f2"
    
    # hapus bagian loading proses
    sheet["D15"].color = None
    sheet["D15"].value = None
    sheet["D16"].value = None
    
    # ambil data path, sheet, dan nama kolom input user
    path = sheet["D11"].value
    nama_sheet = sheet["D12"].value
    nama_patokan = sheet["D13"].value
    
    list_patokan = []
    list_patokan_temp = list(nama_patokan.split(","))
    for patokan in list_patokan_temp:
        list_patokan.append(patokan.strip())
    
    # Ambil data
    wb1 = getData3(path, nama_sheet, sheet, list_patokan)
    
    kolom_terakhir = wb1["A1"].expand("right").last_cell.column
    
    rows = []
    for i in range(1,kolom_terakhir):
        last_row = wb1.range((1, i)).expand("table").last_cell.row
        rows.append(last_row)
    lokasi_kolom = rows.index(max(rows))
    
    sheet["D15"].value = "Process..."
    sheet["D15"].font.color = "#000000"
    sheet["D15"].color = "#f08035"
            
    # ubah data menjadi pandas dataframe
    df = wb1.range((1, lokasi_kolom+1)).options(pd.DataFrame, header=1, index=False, expand='table', empty=np.NaN).value
    df = df.replace(r'^\s*$', np.nan, regex=True)
    df = df.fillna(value=np.nan)
    
    # tambah tulisan 50%
    sheet["D15"].value = "Process 50%..."
    sheet["D15"].font.color = "#000000"
    sheet["D15"].color = "#f08035"
    
    # Ambil Panjang Awal
    panjang_awal = df.shape[0]
    
   #  if len(list_patokan) == 1:
   #    list_patokan = list_patokan[0]
    
    # dropping ALL duplicate values
    df.drop_duplicates(subset=list_patokan, keep="first", inplace=True, ignore_index=True)
   #  df2 = df.drop(df.columns[[0]], axis = 1)
    
    wb1.clear_contents()
    
    wb1.range('A1').options(index=False).value = df
    
    last_row = history.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
    last_value = history.range((last_row,1)).options(index=False).value
    if last_value == "No":
        no = 1
    else:
        no = int(last_value)+1
    
    end = time.time()
    
    path_list = path.split("\\")
    
    history.range((last_row+1,1)).options(index=False, header=False).value = pd.DataFrame([[no, 
                                                                                            dt_string,
                                                                                            path_list[len(path_list)-1],
                                                                                            "Hapus Data Duplikat", 
                                                                                            str(panjang_awal) + " Data", 
                                                                                            "-", 
                                                                                            str(round((end-start)/60, 2)) + " Menit", 
                                                                                            str(panjang_awal - df.shape[0]) + " Data Terhapus"]])
    sheet["D15"].value = "Done!"
    sheet["D15"].font.color = "#000000"
    sheet["D15"].color = "#84de1d"
    sheet["D16"].value = str(panjang_awal - df.shape[0]) + " Data Terhapus"
    
if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('topicDetection.xlsm').set_mock_caller()
    main()


