# -*- coding: utf-8 -*-
"""
Created on Thu Aug 15 22:22:46 2019

@author: Asus
"""
#Import Library
import xlrd#library untuk read file excel
import xlsxwriter#library untuk buat file excel
import math#libary math untuk menggunakan sqrt (akar)
import numpy as np#library digunakan untuk membuat ndarray (type data untuk dimasukan sebagai inisisasi centroid awal)
from sklearn.cluster import KMeans#library untuk menggunakan tools kmeans

def LoadDataset(FileLoc):#fungsi untuk load dataset
    kecamatan = []
    data = []
    workbook = xlrd.open_workbook(FileLoc)
    sheet = workbook.sheet_by_index(0)
    count = 0
    for i in range(1,sheet.nrows):
        kecamatan.append(sheet.cell_value(i,1))
        kabupaten = sheet.cell_value(i,2)
        lintang = sheet.cell_value(i,3)
        bujur = sheet.cell_value(i,4)
        satelit = sheet.cell_value(i,5)
        data.append([kabupaten,lintang,bujur,satelit])
        count += 1
        print(count, "Data loaded")
    return kecamatan,data

def Euclidean(a,b):#fungsi untuk jarak euclide
    distance = 0
    for i in range(len(a)):
        distance = distance + ((a[i]-b[i])**2)
    temp = math.sqrt(distance)
    return round(temp, 3)

def Euclideans(centroid,datas):#fungsi untuk jarak euclide dari 1 centroid ke seluruh data
    c0 = []
    c1 = []
    c2 = []
    for data in datas:
        tempC0 = Euclidean(centroid[0],data)#panggil fungsi jarak euclide diatas
        tempC1 = Euclidean(centroid[1],data)#panggil fungsi jarak euclide diatas
        tempC2 = Euclidean(centroid[2],data)#panggil fungsi jarak euclide diatas
        c0.append(tempC0)
        c1.append(tempC1)
        c2.append(tempC2)
    return c0,c1,c2

kecamatan,data = LoadDataset("data cluster.xls")#panggil fungsi load data
centroid_awal = np.array([[0.294, 0.165, 0.076, 0.500], [0.059, 0.261, 0.023, 0.500], [0.471, 0.070, 0.078, 0.000]])#untuk isi inisisai centroid awal, bisa dipakai bisa gak
kmeans = KMeans(n_clusters=3, random_state=0).fit(data)#proses kmeans menggunakan tools(library) sklearn
labels = kmeans.labels_.tolist()#Hasil label di convert ke list
centroids = kmeans.cluster_centers_.tolist()#Centroid
centroid0 = [round(centroids[0][0],3),round(centroids[0][1],3),round(centroids[0][2],3),round(centroids[0][3],3)]#memisahkan centroid ke list baru
centroid1 = [round(centroids[1][0],3),round(centroids[1][1],3),round(centroids[1][2],3),round(centroids[1][3],3)]#memisahkan centroid ke list baru
centroid2 = [round(centroids[2][0],3),round(centroids[2][1],3),round(centroids[2][2],3),round(centroids[2][3],3)]#memisahkan centroid ke list baru
iteration = kmeans.n_iter_#Jumlah iterasi
distanceC0,distanceC1,distanceC2 = Euclideans(centroids,data)#panggul fungsi untuk hitung euclide dari centroid ke seluruh data

#dari sini kebawah untuk buat file excel
book = xlsxwriter.Workbook("Result.xlsx")
sheet = book.add_worksheet()

sheet.write("H1","K")
sheet.write("I1","L")
sheet.write("J1","B")
sheet.write("K1","S")

sheet.write("L1","K")
sheet.write("M1","L")
sheet.write("N1","B")
sheet.write("O1","S")

sheet.write("P1","K")
sheet.write("Q1","L")
sheet.write("R1","B")
sheet.write("S1","S")

sheet.merge_range("H2:K2","Centroid 0")
sheet.merge_range("L2:O2","Centroid 1")
sheet.merge_range("P2:S2","Centroid 2")

sheet.write("A3","ID.")
sheet.write("B3","Kecamatan")
sheet.write("C3","Kabupaten")
sheet.write("D3","Lintang")
sheet.write("E3","Bujur")
sheet.write("F3","Satelit")
sheet.write("H3",centroid2[0])
sheet.write("I3",centroid2[1])
sheet.write("J3",centroid2[2])
sheet.write("K3",centroid2[3])
sheet.write("L3",centroid1[0])
sheet.write("M3",centroid1[1])
sheet.write("N3",centroid1[2])
sheet.write("O3",centroid1[3])
sheet.write("P3",centroid0[0])
sheet.write("Q3",centroid0[1])
sheet.write("R3",centroid0[2])
sheet.write("S3",centroid0[3])
sheet.write("T3","Centroid 0")
sheet.write("U3","Centroid 1")
sheet.write("V3","Centroid 2")

for i in range(len(kecamatan)):
    line = i + 4
    Id = i + 1
    cellA = "A" + str(line)
    cellB = "B" + str(line)
    cellC = "C" + str(line)
    cellD = "D" + str(line)
    cellE = "E" + str(line)
    cellF = "F" + str(line)
    cellHK = "H" + str(line) + ":" + "K" + str(line)
    cellLO = "L" + str(line) + ":" + "O" + str(line)
    cellPS = "P" + str(line) + ":" + "S" + str(line)
    cellT = "T" + str(line)
    cellU = "U" + str(line)
    cellV = "V" + str(line)
    sheet.write(cellA,Id)
    sheet.write(cellB,kecamatan[i])
    sheet.write(cellC,data[i][0])
    sheet.write(cellD,data[i][1])
    sheet.write(cellE,data[i][2])
    sheet.write(cellF,data[i][3])
    sheet.merge_range(cellHK,distanceC2[i])
    sheet.merge_range(cellLO,distanceC1[i])
    sheet.merge_range(cellPS,distanceC0[i])
    if labels[i] == 2:
        sheet.write(cellT,0)
    elif labels[i] == 1:
        sheet.write(cellU,1)
    else:
        sheet.write(cellV,2)
    
book.close()