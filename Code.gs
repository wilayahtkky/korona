const scriptProp = PropertiesService.getScriptProperties();

function initialSetup() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', doc.getId());
}

function sanitizeValue(value) {
  if (typeof value !== 'string') return value;
  const formulaTriggers = ['=', '+', '-', '@'];
  if (formulaTriggers.some(trigger => value.startsWith(trigger))) {
    return "'" + value;
  }
  return value;
}

function sendAdminNotification(nextRow, data) {
  const recipient = 'wilayahtkky@gmail.com';
  const senderName = 'wil TKKY - Buntoro';
  const subject = 'New Registration: Petugas Perarakan Korona Adven 2026';

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <p>A new submission has been added to row <strong>${nextRow}</strong>.</p>
      <ul style="background-color: #f8f9fa; padding: 15px 25px; border-radius: 6px; list-style-type: none;">
        <li><strong>Wilayah:</strong> ${data['Wilayah'] || '-'}</li>
        <li><strong>Lingkungan:</strong> ${data['Lingkungan'] || '-'}</li>
        <li><strong>Nama Pembawa Korona:</strong> ${data['Nama Pembawa Korona'] || '-'}</li>
        <li><strong>No. HP/WA:</strong> ${data['No. HP/WA Pembawa Korona'] || '-'}</li>
        <li><strong>Email:</strong> ${data['email'] || '-'}</li>
      </ul>
    </div>
  `;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName,
  });
}

function sendUserConfirmationEmail(userEmail, data) {
  const senderName = 'Panitia Natal 2026 - Wilayah Theresia dari Kanak-Kanak Yesus';
  const subject = 'Bukti Pendaftaran Petugas Perarakan Korona Adven 2026';

  // Jadwal image: replace GITHUB_RAW_URL with your actual GitHub raw URL after deploy
  // Example: https://raw.githubusercontent.com/USERNAME/REPO/main/jadwal.jpg
  const jadwalImageUrl = 'https://raw.githubusercontent.com/wilayahtkky/korona/main/jadwal.png';

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px;">
      <h2 style="color: #2d6a4f;">Bukti Pendaftaran Petugas Perarakan Korona Adven 2026</h2>
      <p>Syalom, terima kasih telah melakukan pendaftaran. Berikut adalah rincian data yang telah dikirimkan:</p>
      
      <ul style="background-color: #f8f9fa; padding: 15px 25px; border-radius: 6px; list-style-type: none;">
        <li><strong>Wilayah:</strong> ${data['Wilayah'] || '-'}</li>
        <li><strong>Lingkungan:</strong> ${data['Lingkungan'] || '-'}</li>
        <li><strong>Nama Pembawa Korona:</strong> ${data['Nama Pembawa Korona'] || '-'}</li>
        <li><strong>No. HP/WA Pembawa Korona:</strong> ${data['No. HP/WA Pembawa Korona'] || '-'}</li>
      </ul>

      <p><strong>Pengingat Jadwal Perarakan Korona Adven 2026:</strong></p>
      <ul style="padding-left: 20px;">
        <li>Sabtu, 28 November 2026 = 16.00 &amp; 18.00</li>
        <li>Minggu, 29 November 2026 = 08.30, 11.00, 16.30 &amp; 19.00</li>
        <li><strong>Gladi Bersih</strong> = Minggu 22 November 2026 Pk.13.00 <em>(Wajib hadir)</em></li>
      </ul>

      <p>Pembagian tugas dapat dilihat pada gambar jadwal di bawah. Pastikan nomor WA aktif, kami akan mengingatkan lewat WhatsApp.</p>

      <p style="text-align:center;">
        <img src="${jadwalImageUrl}" alt="Jadwal Perarakan Korona Adven 2026" style="max-width:100%; height:auto; border-radius:8px; border:1px solid #ddd;" />
      </p>

      <p><em>Pembelian Lilin Korona dapat dilakukan saat Sosialisasi Bulan Keluarga Jumat, 20 November 2026. Perwakilan lingkungan yang hadir diharapkan membawa uang pas. Informasi biaya lilin menyusul.</em></p>
      
      <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
      <p style="font-size: 12px; color: #777;">Email ini dikirim secara otomatis. Mohon tidak membalas email ini.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: userEmail,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName,
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  // Coba dapatkan kunci antrian maks 10.000 ms (10 detik)
  const hasLock = lock.tryLock(10000);

  // Jika gagal mendapatkan kunci (antrian penuh / timeout)
  if (!hasLock) {
    return ContentService.createTextOutput(
      JSON.stringify({
        result: 'busy',
        message: 'Server sedang sibuk karena banyaknya pendaftaran bersamaan. Silakan coba klik Kirim Pendaftaran beberapa saat lagi.'
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    // Bot honeypot check
    if (e.parameter.mobile_number && e.parameter.mobile_number !== '') {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'success', message: 'Bot detected' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    let sheet = doc.getSheetByName(e.parameter.sheet_name || 'pembawakorona');

    // Fallback if sheet tab is not found
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextRow = sheet.getLastRow() + 1;

    // --- Duplicate check: one Lingkungan can only register once ---
    const submittedLingkungan = (e.parameter['Lingkungan'] || '').toString().trim();
    const submittedWilayah = (e.parameter['Wilayah'] || '').toString().trim();

    if (submittedLingkungan && sheet.getLastRow() > 1) {
      const idxLingkungan = headers.indexOf('Lingkungan');
      const idxWilayah = headers.indexOf('Wilayah');
      const idxNama = headers.indexOf('Nama Pembawa Korona');

      if (idxLingkungan >= 0) {
        const allData = sheet.getDataRange().getValues();

        for (let i = 1; i < allData.length; i++) {
          const lingk = (allData[i][idxLingkungan] || '').toString().trim();
          const wil = idxWilayah >= 0 ? (allData[i][idxWilayah] || '').toString().trim() : '';
          const namaTerdaftar = idxNama >= 0 ? (allData[i][idxNama] || '').toString().trim() : '';

          if (lingk === submittedLingkungan) {
            if (!submittedWilayah || !wil || wil === submittedWilayah) {
              const namaInfo = namaTerdaftar
                ? ' atas nama "' + namaTerdaftar + '"'
                : '';
              return ContentService.createTextOutput(
                JSON.stringify({
                  result: 'duplicate',
                  message: 'Lingkungan "' + submittedLingkungan + '" sudah terdaftar' + namaInfo + '.\n\nSetiap lingkungan hanya boleh mengirim 1 perwakilan.'
                })
              ).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }
      }
    }

    const newRow = headers.map(function (header) {
      if (header === 'id') return Utilities.getUuid();
      if (header === 'timestamp') return new Date();

      const rawValue = e.parameter[header] || '';
      return sanitizeValue(rawValue);
    });

    const newRange = sheet.getRange(nextRow, 1, 1, newRow.length);
    newRange.setNumberFormat('@');
    newRange.setValues([newRow]);

    // Send admin notification
    sendAdminNotification(nextRow, e.parameter);

    // Send confirmation email to user IF email field is present and non-empty
    const userEmail = e.parameter.email || e.parameter.Email;
    if (userEmail && userEmail.trim() !== '') {
      sendUserConfirmationEmail(userEmail.trim(), e.parameter);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ result: 'success', row: nextRow })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    // Selalu lepaskan lock setelah proses selesai
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    let sheet = doc.getSheetByName('pembawakorona');

    if (!sheet) {
      sheet = doc.getSheets()[0];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return responseOutput([], e);
    }

    const headers = data[0];
    const rows = data.slice(1);

    // Cari indeks kolom
    const idxWilayah = headers.indexOf('Wilayah');
    const idxLingkungan = headers.indexOf('Lingkungan');
    const idxNama = headers.indexOf('Nama Pembawa Korona');

    // Hanya ambil kolom publik yang dibutuhkan laporan
    const publicData = rows.map(row => {
      return {
        'Wilayah': idxWilayah >= 0 ? row[idxWilayah] : '',
        'Lingkungan': idxLingkungan >= 0 ? row[idxLingkungan] : '',
        'Nama Pembawa Korona': idxNama >= 0 ? row[idxNama] : ''
      };
    });

    return responseOutput(publicData, e);

  } catch (err) {
    return responseOutput({ error: err.toString() }, e);
  }
}

// Fungsi pembantu untuk memformat output JSON / JSONP
function responseOutput(data, e) {
  const callback = e && e.parameter ? (e.parameter.callback || e.parameter.prefix) : null;
  
  if (callback) {
    const output = callback + '(' + JSON.stringify(data) + ');';
    return ContentService.createTextOutput(output)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
