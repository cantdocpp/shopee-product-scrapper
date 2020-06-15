var xl = require('excel4node');

// Create a new instance of a Workbook class
var wb = new xl.Workbook();
 
// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');

ws.cell(1, 1)
    .string('Kategori')

ws.cell(1, 2)
    .string('Nama Produk')

ws.cell(1, 3)
    .string('Deskripsi Produk')

ws.cell(1, 4)
    .string('SKU Induk')

ws.cell(1, 5)
    .string('Kode Integrasi Variasi')

ws.cell(1, 6)
    .string('Nama Variasi 1')

ws.cell(1, 7)
    .string('Varian untuk Variasi 1')

ws.cell(1, 8)
    .string('Foto produk Per Varian')

ws.cell(1, 9)
    .string('Nama Variasi 2')

ws.cell(1, 10)
    .string('Varian untuk Variasi 2')

ws.cell(1, 11)
    .string('Harga')

ws.cell(1, 12)
    .string('Stok')

ws.cell(1, 13)
    .string('Kode Variasi')

ws.cell(1, 14)
    .string('Foto Sampul')

ws.cell(1, 15)
    .string('Foto Produk 1')

ws.cell(1, 16)
    .string('Foto Produk 2')

ws.cell(1, 17)
    .string('Foto Produk 3')

ws.cell(1, 18)
    .string('Foto Produk 4')

ws.cell(1, 19)
    .string('Foto Produk 5')

ws.cell(1, 20)
    .string('Foto Produk 6')

ws.cell(1, 21)
    .string('Foto Produk 7')

ws.cell(1, 22)
    .string('Foto Produk 8')

ws.cell(1, 23)
    .string('Berat')

ws.cell(1, 24)
    .string('Panjang')

ws.cell(1, 25)
    .string('Lebar')

ws.cell(1, 26)
    .string('Tinggi')

ws.cell(1, 27)
    .string('J&T Express')

ws.cell(1, 28)
    .string('Sicepat REG')

ws.cell(1, 29)
    .string('Sicepat Halu')

ws.cell(1, 30)
    .string('JNE Reguler (Cashless)')

ws.cell(1, 31)
    .string('JNE JTR (Cashless)')

ws.cell(1, 32)
    .string('JNE YES (Cashless)')

ws.cell(1, 33)
    .string('GoSend Same Day')

ws.cell(1, 34)
    .string('GrabExpress Sameday')

ws.cell(1, 35)
    .string('Dikirim Dalam Pre-order')
    

module.exports = {
    ws,
    wb
}
