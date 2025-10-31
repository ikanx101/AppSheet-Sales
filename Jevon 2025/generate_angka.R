# Skrip untuk mengenerate angka antara elemen-elemen dalam object marker_1

# Data marker_1 (berdasarkan informasi dari environment)
marker_1 <- c(8, 23, 31, 44, 52, 60, 85, 92, 103, 114, 135, 143, 155, 163, 190, 198, 208, 219, 232, 241, 248, 255, 268, 276, 287, 318, 337, 344)

# Fungsi untuk generate angka antara dua nilai
generate_angka_antara <- function(start, end) {
  if (end - start <= 1) {
    return(NULL)
  }
  seq(from = start + 1, to = end - 1, by = 1)
}

# Generate angka untuk setiap pasangan berurutan
hasil_generate <- list()

for (i in 1:(length(marker_1) - 1)) {
  start <- marker_1[i]
  end <- marker_1[i + 1]
  
  # Generate angka antara start dan end
  angka_antara <- generate_angka_antara(start, end)
  
  if (!is.null(angka_antara)) {
    hasil_generate[[paste0("marker[", i, "] - marker[", i + 1, "]")]] <- angka_antara
  }
}

# Tampilkan hasil
cat("Object marker_1:")
print(marker_1)
cat("\n\nHasil generate angka antara elemen-elemen marker_1:\n\n")

for (nama in names(hasil_generate)) {
  cat(nama, ": ", hasil_generate[[nama]], "\n")
}

# Alternatif: Simpan semua angka dalam satu vector
semua_angka <- unlist(hasil_generate)
cat("\n\nSemua angka yang digenerate (dalam satu vector):\n")
print(semua_angka)

# Menghitung jumlah angka yang digenerate
cat("\nTotal angka yang digenerate:", length(semua_angka))

# Contoh penggunaan untuk kasus spesifik
cat("\n\nContoh untuk beberapa range pertama:")
cat("\nAngka antara marker[1] (8) dan marker[2] (23):", generate_angka_antara(8, 23))
cat("\nAngka antara marker[2] (23) dan marker[3] (31):", generate_angka_antara(23, 31))
cat("\nAngka antara marker[3] (31) dan marker[4] (44):", generate_angka_antara(31, 44))
