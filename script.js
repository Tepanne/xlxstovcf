document.getElementById('file-upload').addEventListener('change', handleFile, false);
document.getElementById('convert-btn').addEventListener('click', convertToVCF, false);

let excelData = null;
let sheetNames = [];

// Fungsi untuk membaca file Excel
function handleFile(event) {
    const reader = new FileReader();
    const file = event.target.files[0];

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Simpan nama sheet
        sheetNames = workbook.SheetNames;

        // Tambahkan nama sheet ke dropdown
        const sheetSelect = document.getElementById('sheet-select');
        sheetSelect.innerHTML = '';
        sheetNames.forEach((name, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        // Pilih sheet pertama secara default
        sheetSelect.selectedIndex = 0;

        // Tampilkan isi sheet pertama
        displaySheet(workbook, sheetNames[0]);

        // Tambah event listener untuk memilih sheet
        sheetSelect.addEventListener('change', () => {
            displaySheet(workbook, sheetNames[sheetSelect.value]);
        });
    };

    reader.readAsArrayBuffer(file);
}

// Fungsi untuk menampilkan isi sheet dalam bentuk tabel
function displaySheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    const table = document.getElementById('excel-table');
    table.innerHTML = ''; // Hapus konten tabel sebelumnya

    // Menambahkan header tabel
    const headerRow = table.insertRow(-1);
    headerRow.className = 'header-row';

    // Tambahkan penanda kolom
    const headerCell = document.createElement('th');
    headerCell.className = 'header-column';
    headerRow.appendChild(headerCell);

    excelData[0].forEach((header, index) => {
        const cell = headerRow.insertCell(-1);
        cell.textContent = String.fromCharCode(65 + index); // Konversi index ke huruf (A, B, C, ... )
    });

    // Menambahkan isi tabel
    for (let i = 0; i < excelData.length; i++) {
        const row = table.insertRow(-1);
        const rowHeaderCell = row.insertCell(-1);
        rowHeaderCell.textContent = i + 1;

        excelData[i].forEach((cellData, index) => {
            const cell = row.insertCell(-1);
            cell.textContent = cellData || '';
        });
    }
}

// Fungsi untuk parsing rentang input
function parseRange(range) {
    const [start, end] = range.split(':');
    const startColumn = start.match(/[A-Za-z]+/)[0].toUpperCase();
    const startRow = parseInt(start.match(/\d+/)[0], 10);
    const endRow = end ? parseInt(end.match(/\d+/)[0], 10) : null;
    const colIndex = startColumn.charCodeAt(0) - 65;

    return { startRow: startRow - 1, endRow: endRow ? endRow - 1 : null, colIndex };
}

// Fungsi untuk mengonversi ke VCF dan otomatis mengunduh file
// Fungsi untuk mengonversi ke VCF dan otomatis mengunduh file
function convertToVCF() {
    if (!excelData) {
        alert("Silakan unggah file Excel terlebih dahulu.");
        return;
    }

    // Ambil rentang kontak admin, navy, dan anggota
    const adminRange = document.getElementById('admin-range').value;
    const navyRange = document.getElementById('navy-range').value;
    const memberRange = document.getElementById('member-range').value;
    const adminPrefix = document.getElementById('admin-prefix').value || 'Admin';
    const navyPrefix = document.getElementById('navy-prefix').value || 'Navy';
    const memberPrefix = document.getElementById('member-prefix').value || 'Anggota';
    const outputFilename = document.getElementById('output-filename').value || 'Kontak';

    let adminVCFContent = "";
    let navyVCFContent = "";
    let memberVCFContent = "";
    let hasAdminContacts = false;
    let hasNavyContacts = false;

    // Proses kontak admin
    if (adminRange) {
        const { startRow, endRow, colIndex } = parseRange(adminRange);
        let adminCounter = 1; // Mulai dari 1 untuk setiap klasifikasi
        for (let i = startRow; i <= (endRow !== null ? endRow : excelData.length - 1); i++) {
            let phone = excelData[i][colIndex];
            if (phone) {
                phone = String(phone);
                if (!phone.startsWith('+')) {
                    phone = '+' + phone;
                }
                const contactName = `${adminPrefix} ${adminCounter++}`; // Nomor urut untuk Admin
                adminVCFContent += `BEGIN:VCARD\nVERSION:3.0\nFN:${contactName}\nTEL:${phone}\nEND:VCARD\n\n`;
                hasAdminContacts = true; // Menandai bahwa ada kontak admin
            }
        }
    }

    // Proses kontak navy
    if (navyRange) {
        const { startRow, endRow, colIndex } = parseRange(navyRange);
        let navyCounter = 1; // Mulai dari 1 untuk setiap klasifikasi
        for (let i = startRow; i <= (endRow !== null ? endRow : excelData.length - 1); i++) {
            let phone = excelData[i][colIndex];
            if (phone) {
                phone = String(phone);
                if (!phone.startsWith('+')) {
                    phone = '+' + phone;
                }
                const contactName = `${navyPrefix} ${navyCounter++}`; // Nomor urut untuk Navy
                navyVCFContent += `BEGIN:VCARD\nVERSION:3.0\nFN:${contactName}\nTEL:${phone}\nEND:VCARD\n\n`;
                hasNavyContacts = true; // Menandai bahwa ada kontak navy
            }
        }
    }

    // Proses kontak anggota
    if (memberRange) {
        const { startRow, endRow, colIndex } = parseRange(memberRange);
        let memberCounter = 1; // Mulai dari 1 untuk setiap klasifikasi
        for (let i = startRow; i <= (endRow !== null ? endRow : excelData.length - 1); i++) {
            let phone = excelData[i][colIndex];
            if (phone) {
                phone = String(phone);
                if (!phone.startsWith('+')) {
                    phone = '+' + phone;
                }
                const contactName = `${memberPrefix} ${memberCounter++}`; // Nomor urut untuk Anggota
                memberVCFContent += `BEGIN:VCARD\nVERSION:3.0\nFN:${contactName}\nTEL:${phone}\nEND:VCARD\n\n`;
            }
        }
    }

    // Cek apakah perlu memisahkan file Admin + Navy dan Anggota
    const splitFiles = document.getElementById('split-files').checked;
    if (splitFiles) {
        let adminNavyFilename = ""; // Nama file Admin + Navy yang sesuai
        if (hasAdminContacts && hasNavyContacts) {
            adminNavyFilename = `Admin_Navy_${outputFilename}`;
        } else if (hasAdminContacts) {
            adminNavyFilename = `Admin_${outputFilename}`;
        } else if (hasNavyContacts) {
            adminNavyFilename = `Navy_${outputFilename}`;
        }

        // Jika ada kontak Admin atau Navy, unduh file tersebut
        if (adminVCFContent || navyVCFContent) {
            downloadFile(adminVCFContent + navyVCFContent, `${adminNavyFilename}.vcf`);
        }

        // Unduh file Anggota jika ada (gunakan hanya nama file tanpa prefix)
        if (memberVCFContent) {
            downloadFile(memberVCFContent, `${outputFilename}.vcf`);
        }
    } else {
        // Gabungkan semua kontak jika tidak dipisahkan
        const combinedVCFContent = adminVCFContent + navyVCFContent + memberVCFContent;
        downloadFile(combinedVCFContent, `${outputFilename}.vcf`);
    }
}

// Fungsi untuk mengunduh file VCF
function downloadFile(content, filename) {
    const blob = new Blob([content], { type: 'text/vcard' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
