let excelData = [];
let staffNames = [];
let headers = [];
let rows = [];

// Read Excel file
document.getElementById("excelFile").addEventListener("change", function (e) {
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert to JSON with header row
    excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract headers and data
    headers = excelData[0];
    rows = excelData.slice(1);

    // Get unique staff names from column index 22 (Staff)
    staffNames = [...new Set(rows.map((row) => row[22]))].filter(
      (name) => name
    );

    if (staffNames.length > 0) {
      displayStaffSelection(staffNames);
      document.getElementById("staffSelection").classList.remove("hidden");
    } else {
      alert("No staff members found in the Excel file!");
    }
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});

// Display staff selection radio buttons
function displayStaffSelection(staffNames) {
  const radioGroup = document.getElementById("staffRadioGroup");
  radioGroup.innerHTML = "";

  staffNames.forEach((staff, index) => {
    const div = document.createElement("div");
    div.className = "staff-radio";

    const radio = document.createElement("input");
    radio.type = "radio";
    radio.id = `staff-${index}`;
    radio.name = "staff";
    radio.value = staff;
    if (index === 0) radio.checked = true;
    radio.addEventListener("change", () => displayStaffData(staff));

    const label = document.createElement("label");
    label.htmlFor = `staff-${index}`;
    label.textContent = staff;

    div.appendChild(radio);
    div.appendChild(label);
    radioGroup.appendChild(div);
  });

  // Display data for the first staff by default
  displayStaffData(staffNames[0]);
}

// Display data for selected staff (using indexes)
function displayStaffData(staffName) {
  const filteredData = rows.filter((row) => row[22] === staffName);

  document.getElementById("selectedStaffName").textContent = staffName;
  const tableBody = document.getElementById("recordsTableBody");
  tableBody.innerHTML = "";

  filteredData.forEach((row) => {
    const tr = document.createElement("tr");

    const tdName = document.createElement("td");
    tdName.textContent = row[16] || "N/A";

    const tdConsult = document.createElement("td");
    tdConsult.textContent = row[11] || "N/A";

    const tdDate = document.createElement("td");
    let completionDate = "N/A";
    if (row[7]) {
      if (typeof row[7] === "number") {
        const date = excelDateToJSDate(row[7]);
        completionDate = date.toLocaleDateString();
      } else {
        const date = new Date(row[7]);
        completionDate = isNaN(date.getTime())
          ? row[7]
          : date.toLocaleDateString();
      }
    }
    tdDate.textContent = completionDate;

    const tdAction = document.createElement("td");
    const button = document.createElement("button");
    button.className = "generate-btn";
    button.textContent = "Generate PDF";
    button.addEventListener("click", () => generatePDF(row));
    tdAction.appendChild(button);

    tr.appendChild(tdName);
    tr.appendChild(tdConsult);
    tr.appendChild(tdDate);
    tr.appendChild(tdAction);

    tableBody.appendChild(tr);
  });

  document.getElementById("dataDisplay").classList.remove("hidden");
}

// Helper for wrapping text + new page
function addWrappedText(
  doc,
  text,
  x,
  y,
  maxWidth,
  lineHeight,
  bottomMargin = 280
) {
  const lines = doc.splitTextToSize(text, maxWidth);
  let cursorY = y;

  lines.forEach((line) => {
    if (cursorY > bottomMargin) {
      doc.addPage();
      cursorY = 30; // reset top margin
    }
    doc.text(line, x, cursorY);
    cursorY += lineHeight;
  });

  return cursorY;
}

function generatePDF(row) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });

  // Extract fields
  const rujTuan = row[5] || "N/A";
  const consultant = row[11] || "N/A";
  const alamatconsultant = row[12] || "N/A";
  const jawatantetuan = row[19] || "N/A";
  const projek = row[10] || "N/A";
  const tapak = row[9] || "N/A";
  const namaTetuan = row[16] || "N/A";

  // Handle date
  let completionDate = "N/A";
  if (row[7]) {
    if (typeof row[7] === "number") {
      completionDate = excelDateToJSDate(row[7]).toLocaleDateString("ms-MY");
    } else {
      const d = new Date(row[7]);
      completionDate = isNaN(d.getTime())
        ? row[7]
        : d.toLocaleDateString("ms-MY");
    }
  }

  // --- Header ---
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.text("Ruj. Tuan: " + rujTuan, 20, 30);
  doc.text("Tarikh: " + completionDate, 20, 35);

  // --- Address ---
  doc.setFont("helvetica", "bold");
  doc.text(consultant, 20, 55);
  doc.setFont("helvetica", "normal");
  let y = addWrappedText(doc, alamatconsultant, 20, 60, 85, 7);

  y = addWrappedText(doc, "U.P : " + jawatantetuan, 20, y + 6, 165, 7);

  // --- Subject ---
  y = addWrappedText(doc, "Tuan/Puan,", 20, y + 6, 165, 7);
  doc.setFont("helvetica", "bold");
  y = addWrappedText(doc, "PERMOHONAN KEBENARAN MERANCANG", 20, y + 2, 165, 7);

  doc.setFont("helvetica", "normal");
  y = addWrappedText(doc, "Projek: " + projek, 20, y + 1, 165, 7);
  y = addWrappedText(
    doc,
    "Hartanah/Tapak Projek: " + tapak,
    20,
    y + 1,
    165,
    7
  );
  y = addWrappedText(doc, "Pemilik/Pemaju: " + namaTetuan, 20, y + 1, 165, 7);

  // --- Body ---
  y = addWrappedText(
    doc,
    "Per: Sokongan Merancang Pembangunan",
    20,
    y + 3,
    165,
    7
  );
  doc.line(20, y, 190, y); // underline
  y += 10;

  const body = `
Dengan segala hormatnya saya merujuk kepada perkara tersebut di atas dan surat Arkitek/Jurutera Perunding lantikan tuan bertarikh 6 OGOS 2025 berhubung perkara yang sama.
2.  Sukacita dimaklumkan bahawa Telekom Malaysia Berhad (“TM”) bersedia untuk memberi sokongan
kepada pembangunan yang dirancang bagi Projek tersebut di atas tertakluk kepada terma-terma dan syarat-syarat berikut:
    (a) pemilik/pemaju menyatakan tarikh mula dan jangka siap Projek tersebut; 
    (b) pemilik/pemaju melantik perunding berdaftar untuk menyedia dan merekabentuk insfrastruktur rangkaian telekomukikasi talian dari dan ke Tapak Projek;
    (c) mengemukakan tiga (3) set pelan susunatur yang mengandungi susunatur sivil, lurang/manholes, internal trunking dan bilik telekomunikasi/Subscriber Distribution Frame/Fiber Room (“Bilik Telekomunikasi”) berserta diagram skematik rangkaian telekomunikasi talian tetap. Pelan yang dikemukakan hendaklah mengikut spesifikasi yang telah ditetapkan oleh Suruhanjaya Komunikasi dan Multimedia Malaysia (“SKMM”) atau Malaysian Technical Standards Forum Bhd (“MTSFB”);
    (d) pemaju/pemilik bersetuju tidak mengenakan sebarang fi atau caj ke atas TM untuk penempatan infrastruktur telekomunikasi bagi perkhidmatan telekomunikasi yang diberikan oleh TM kepada pelanggan-pelanggannya di bangunan berkenaan sepanjang tempoh perkhidmatan tersebut disediakan;  
    (e) menyediakan pendawaian meter elektrik berasingan bagi bekalan tenaga elektrik untuk kegunaan peralatan telekomunikasi;
    (f) mendapatkan kelulusan jajaran daripada Jabatan Bangunan pihak berkuasa tempatan, jika berkenaan (bagi keperluan infrastruktur untuk status MSC); 
    (g) menyediakan pendawaian dalaman jenis gentian fiber optik single mode (bagi kawasan High Speed Broadband “HSBB” & tertakluk kepada jenis pembangunan); 
    (h) Sila Ambil Maklum bahawa pihak TM akan mengenakan fi pemprosesan ke atas kerja-kerja semakan dan kelulusan pelan infrastruktur telekomunikasi projek yang tersebut di atas bergantung kepada jenis pembangunan yang dirancang. Fi yang dikenakan terbahagi kepada 2 peringkat seperti jadual-jadual harga di bawah: 
        Peringkat 1: Jadual harga kerja-kerja semakan dan kelulusan pelan
        Peringkat 2: Jadual harga kerja-kerja ujian penerimaan bagi tujuan pengeluaran 'Certificate pf Acceptance' (COA) dan Pengesahan Kelulusan Siap Insfrastruktur Telekomunikasi
        Nota: Fi pemprosesan di atas tertakluk kepada Cukai Jualan dan Perkhidmatan (SST) sebanyak 8%, dan dikecualikan bagi Zon Bebas Cukai (Langkawi, Labuan and Tioman)
    (i) sekiranya terdapat perubahan besar (“major changes”) terhadap pelan asal Projek tersebut dan pihak tuan mengemukakan semula pelan pembangunan kepada TM untuk tujuan semakan dan kelulusan, pihak TM akan mengenakan fi pemprosesan bagi kerja-kerja semakan dan kelulusan semula seperti jadual harga berikut:
        Nota: “Major Changes” adalah merujuk kepada perubahan pelan disebabkan oleh perubahan rekabentuk pembangunan yang melibatkan semakan semula rekabentuk sivil, gentian optik dalaman dan kelulusan semula pelan telekomunikasi. Fi pemprosesan semula di atas tertakluk kepada Cukai Perkhidmatan pada kadar 8%, dan dikecualikan bagi Zon Bebas Cukai (Langkawi, Labuan and Tioman) 
    (j) segala kerja-kerja mestilah dijalankan oleh kontraktor kerja awam yang mempunyai:
             i.   Sijil sah PKK Kepala VIII – Kerja-Kerja Telekomunikasi; atau 
             ii.  Sijil sah CIDB di bawah kod E08 dan CE21
        Pihak TM akan menyelia dan menyemak kerja-kerja semasa ujian penerimaan infrastruktur telekomunikasi dijalankan;
    (k) pelan yang dikemukakan mestilah mematuhi syarat-syarat seperti yang dinyatakan di dalam “Guideline On The Provision Of Basic Civil Works For Communications Infrastructure In New Development Areas (SKMM/G/01/09)” yang dikeluarkan oleh SKMM;
    (l) terma-terma dan syarat-syarat yang dinyatakan di dalam surat ini adalah mengikat ke atas pewaris, pengganti dalam hak milik, wasi, pentadbir semasa, wakil peribadi atau penerima serahan hak pemilik/pemaju.
3.  Sila Ambil Maklum bahawa sekiranya terdapat infrastruktur telekomunikasi sediada TM (“Infrastruktur Telekomunikasi TM”) di dalam kawasan projek tersebut, pihak tuan adalah bertanggungjawab sepenuhnya untuk mengalihkan Infrastruktur Telekomunikasi TM mengikut cadangan TM dan menanggung semua kos yang berkaitan dengan kerja-kerja pengalihan Infrastruktur Telekomunikasi TM terbabit.
4.  Sila Ambil Maklum bahawa pihak tuan adalah bertanggungan dan perlu memastikan bahawa kerja-kerja pengalihan Infrastruktur Telekomunikasi TM dijalankan dengan teliti supaya tidak berlaku sebarang kemusnahan atau kerosakan kepada Infrastruktur Telekomunikasi TM di mana sebarang kos baik pulih beserta ganti rugi akan menjadi tanggungjawab pihak tuan.
5.  Sila Ambil Maklum bahawa sekiranya semasa pihak tuan menjalankan projek tersebut di dalam dan/atau di luar kawasan pembangunan dan berlaku kemalangan yang menyebabkan kerosakan terhadap Infrastruktur Telekomunikasi TM, pihak Tuan adalah bertanggungjawab sepenuhnya untuk menanggung segala kos baik pulih berserta bayaran balik kerugian yang ditanggung oleh pihak TM.
6.  Sila Ambil Maklum, bahawa semua pemasangan baru Infrastruktur Telekomunikasi TM (jika berkenaan) adalah muktamad dan tiada perancangan untuk pengalihan di masa hadapan. Namun begitu, jika terdapat sebarang keperluan dari pihak tuan untuk mengalihkan infrastruktur telekomunikasi sedia ada, TM tidak akan bertanggungjawab untuk menanggung sebarang kos pengalihan dan pemasangan semula infrastruktur tersebut. Pihak Tuan hendaklah memastikan cadangan dan perancangan Projek di atas telah mengambilkira keseluruhan perancangan pembangunan di masa hadapan.
7.  Sila Ambil Maklum, pihak pemaju hendaklah memastikan bahawa setiap FTB/DP dan laluan trunking berada di bahagian luar pintu pagar kedai dan elakkan berada di kawasan terkurung.
8.  Sokongan ini tidak mengikat TM untuk memberikan sebarang perkhidmatan telekomunikasi setelah kerja-kerja pembangunan siap dilaksanakan
9.  Sokongan ini hanya diperakui untuk jangkamasa 2 (dua) tahun dari tarikh surat ini. Pelan-pelan perlu dikemukakan semula untuk kelulusan selepas tamat tempoh.
10. Sekiranya pihak tuan bersetuju atau tidak bersetuju dengan jadual harga, terma-terma dan syarat-syarat yang dinyatakan di atas, sila kemukakan Perakuan Persetujuan tuan (Lampiran 1) kepada kami selewat-lewatnya pada 30 September 2025 untuk tindakan kami yang selanjutnya.
11. Sekiranya pihak tuan memerlukan sebarang penjelasan berkaitan dengan perkara tersebut di atas, sila hubungi Shahrin Amirul Bin Ab Rahman (shahrinamirul.abrahman@tm.com.my) di talian 011-10288084 atau Noriza Binti Ahmad Sa’don (noriza.ahmadsadon@tm.com.my) di talian 013-2068582. 

`;

  y = addWrappedText(doc, body, 20, y, 165, 7);

  // --- Closing ---
  if (y > 240) {
    doc.addPage();
    y = 40;
  }

  y = addWrappedText(doc, "Sekian, terima kasih.", 20, y + 10, 165, 7);
  y = addWrappedText(doc, "Yang benar,", 20, y + 10, 165, 7);

  doc.text(".............................................", 20, y + 20);
  doc.text("AIZAN BIN ATIMAN", 20, y + 30);
  doc.text("Penolong Pengurus", 20, y + 40);
  doc.text("Telekom Malaysia Berhad", 20, y + 50);

  // Save
  doc.save("Surat_Penuh_" + namaTetuan.replace(/\s+/g, "_") + ".pdf");
}

// Convert Excel date to JavaScript Date object
function excelDateToJSDate(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);

  const fractional_day = serial - Math.floor(serial);
  let total_seconds = Math.round(86400 * fractional_day);
  const seconds = total_seconds % 60;
  total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60 * 60));
  const minutes = Math.floor((total_seconds - hours * 60 * 60) / 60);

  date_info.setHours(hours);
  date_info.setMinutes(minutes);
  date_info.setSeconds(seconds);

  return date_info;
}
