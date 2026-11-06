# Keluhan Pembeli Sunscreen dari 1,09 Juta Ulasan Skincare Sephora

Proyek ini menganalisis ulasan produk skincare di Sephora untuk menjawab dua pertanyaan: kategori skincare mana yang paling sering mengecewakan pembeli, dan apa yang mereka keluhkan. Hasilnya dipakai untuk menyusun masukan produk bagi brand sunscreen lokal, dengan Wardah dan Emina sebagai contoh.

Hasil akhir tersedia dalam tiga bentuk: workbook Excel dengan dashboard berfilter, tiga notebook Python, dan dashboard HTML yang bisa dibuka tanpa internet.

## Hasil utama

Sunscreen adalah kategori besar dengan proporsi ulasan negatif tertinggi. Dari 41.100 ulasan sunscreen, 13,8% memberi rating 1–2 bintang, sedangkan rata-rata seluruh skincare 10,4%. Self Tanners (16,0%) dan High Tech Tools (15,0%) lebih tinggi, tetapi jumlah ulasannya hanya 12 ribu dan 6 ribu.

Hampir semua keluhan pembeli sunscreen menyangkut pengalaman memakai produk. Keluhan bahwa produk tidak melindungi dari matahari hanya muncul di 2,8% ulasan negatif. Tabel berikut dihitung dari 5.675 ulasan negatif dan 31.934 ulasan positif sunscreen:

| Keluhan | % ulasan negatif | % ulasan positif | Lift |
|---|---|---|---|
| Berminyak, lengket, atau berat | 32,4% | 15,1% | 2,1× |
| Jerawat atau breakout | 15,1% | 4,5% | 3,4× |
| Pilling (menggumpal) | 14,0% | 1,6% | 8,6× |
| White cast | 13,5% | 5,8% | 2,3× |
| Tidak melindungi dari matahari | 2,8% | 0,9% | 3,0× |

Lift adalah persentase di ulasan negatif dibagi persentase di ulasan positif. Pilling jarang disebut oleh pembeli yang puas, sehingga tema ini paling kuat membedakan ulasan negatif. Bau atau wangi disebut di 15% ulasan negatif tetapi lebih sering di ulasan positif (lift 0,7×), jadi aroma bukan penyebab utama kekecewaan.

Keluhan white cast meningkat seiring warna kulit pengulas:

| Warna kulit | Ulasan negatif | Menyebut white cast |
|---|---|---|
| Terang | 3.002 | 11% |
| Medium | 1.422 | 14% |
| Sawo matang | 564 | 23% |
| Gelap | 156 | 39% |

Kelompok gelap hanya punya 156 ulasan negatif, jadi angkanya kurang stabil. Pola kenaikannya tetap konsisten di keempat kelompok.

![Keluhan white cast menurut warna kulit](reports/figures/08_white_cast_by_skin_tone.png)

Temuan lain:

- Harga hampir tidak berhubungan dengan rating produk (korelasi Spearman 0,13). Persentase pengulas yang merekomendasikan produk adalah 81,6% untuk produk di bawah $25, 84,0% untuk $25–50, 85,9% untuk $50–100, dan 83,9% untuk produk di atas $100.
- Pada kulit berminyak, 41% ulasan negatif sunscreen mengeluhkan tekstur berminyak dan 19% mengeluhkan jerawat. Pada kulit kering, keluhan iritasi lebih tinggi (14%, dibanding 9% pada kulit berminyak).
- Di antara sunscreen dengan sedikitnya 1.000 ulasan, proporsi ulasan negatif berkisar dari 1,3% (Shiseido Urban Environment Fresh-Moisture) sampai 33,3% (Drunk Elephant Umbra Sheer).

## Rekomendasi

Rekomendasi diurutkan menurut seberapa sering masalahnya muncul di data.

1. Tekstur ringan dengan hasil akhir matte. Keluhan berminyak menempati urutan pertama di semua jenis kulit.
2. Formula yang tidak meninggalkan white cast pada kulit sawo matang, diuji pada warna kulit konsumen Indonesia.
3. Tidak menggumpal saat dilapisi skincare atau makeup.
4. Klaim non-comedogenic yang didukung pengujian, karena jerawat adalah keluhan kedua terbanyak.
5. Varian yang lebih lembut untuk kulit kering atau sensitif.

Rekomendasi ini berasal dari data pasar Amerika Serikat. Sebelum dipakai untuk keputusan produk, pola yang sama perlu dicek pada ulasan e-commerce Indonesia.

## Cara melihat hasil

| File | Isi |
|---|---|
| `Portofolio_Ulasan_Skincare_Sephora.xlsx` | Dashboard Excel dengan filter jenis kulit, warna kulit, dan tahun. Mulai dari sheet Dashboard, lalu telusuri setiap angka ke sheet Analisis. |
| `dashboard/index.html` | Versi web dari dashboard. Buka langsung di browser, tidak perlu internet. |
| `notebooks/01_data_understanding_cleaning.ipynb` | Pemuatan data, pengecekan kualitas, dan pembersihan. |
| `notebooks/02_eda.ipynb` | Analisis per harga, kategori, jenis kulit, tahun, dan brand. |
| `notebooks/03_sunscreen_complaints.ipynb` | Deteksi tema keluhan sunscreen dan rekomendasi. |
| `reports/figures/` | Sembilan grafik PNG dari notebook. |

Di workbook Excel, sel kuning di sheet Dashboard adalah filter. Semua KPI, grafik, dan heatmap dihitung ulang dengan rumus SUMIFS setiap kali filter diganti.
