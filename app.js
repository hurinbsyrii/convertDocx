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

// Display staff selection checkboxes
function displayStaffSelection(staffNames) {
  const radioGroup = document.getElementById("staffRadioGroup");
  radioGroup.innerHTML = "";

  staffNames.forEach((staff, index) => {
    const div = document.createElement("div");
    div.className = "staff-checkbox";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.id = `staff-${index}`;
    checkbox.name = "staff";
    checkbox.value = staff;

    const label = document.createElement("label");
    label.htmlFor = `staff-${index}`;
    label.textContent = staff;

    div.appendChild(checkbox);
    div.appendChild(label);
    radioGroup.appendChild(div);
  });

  // Add a button to show records for selected staff
  let showBtn = document.getElementById("showStaffRecordsBtn");
  if (!showBtn) {
    showBtn = document.createElement("button");
    showBtn.id = "showStaffRecordsBtn";
    showBtn.textContent = "Show Records";
    showBtn.className = "generate-btn";
    radioGroup.parentElement.appendChild(showBtn);
  }
  showBtn.onclick = function () {
    const checked = Array.from(
      document.querySelectorAll('input[name="staff"]:checked')
    ).map((cb) => cb.value);
    if (checked.length === 0) {
      alert("Please select at least one staff.");
      return;
    }
    displayStaffData(checked);
  };
}

// Display data for selected staff (array of names)
function displayStaffData(staffNamesArr) {
  const filteredData = rows.filter((row) => staffNamesArr.includes(row[22]));

  document.getElementById("selectedStaffName").textContent = staffNamesArr.join(", ");
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
    button.textContent = "Generate DOCX";
    button.addEventListener("click", () => generateDOCX(row));
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

async function generateDOCX(row) {
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    AlignmentType,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    ShadingType,
    VerticalAlign,
  } = window.docx;

  // ========= Excel → fields (from your indexes) =========
  const rujKami = row[4] || ""; // if you have it; else blank
  const rujTuan = row[5] || "N/A"; // F – 5
  const tarikhRaw = row[7] || null; // H – 7
  const tapak = row[9] || "N/A"; // J – 9
  const projek = row[10] || "N/A"; // K – 10
  const consultant = row[11] || "N/A"; // L – 11
  const alamatConsultant = row[12] || "N/A"; // M – 12
  const namaTetuan = row[16] || "N/A"; // Q – 16
  const uP = row[19] || "N/A"; // T – 19

  // ========= Date helpers =========
  function excelToDate(val) {
    if (!val) return null;
    if (typeof val === "number") {
      const utcDays = Math.floor(val - 25569);
      const utcValue = utcDays * 86400;
      const date = new Date(utcValue * 1000);
      const fractional = val - Math.floor(val);
      let totalSec = Math.round(86400 * fractional);
      const sec = totalSec % 60;
      totalSec -= sec;
      const hrs = Math.floor(totalSec / 3600);
      const mins = Math.floor((totalSec - hrs * 3600) / 60);
      date.setHours(hrs, mins, sec, 0);
      return date;
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  const tarikhDate = excelToDate(tarikhRaw);

  const msMonths = [
    "JANUARI",
    "FEBRUARI",
    "MAC",
    "APRIL",
    "MEI",
    "JUN",
    "JULAI",
    "OGOS",
    "SEPTEMBER",
    "OKTOBER",
    "NOVEMBER",
    "DISEMBER",
  ];
  function fmtDateMalayUpper(d) {
    if (!d) return "N/A";
    return `${d.getDate()} ${msMonths[d.getMonth()]} ${d.getFullYear()}`;
  }

  // Due date to mirror screenshot: **end of next month** from Tarikh
  function endOfNextMonth(d) {
    if (!d) return "N/A";
    const next = new Date(d.getFullYear(), d.getMonth() + 2, 0); // day 0 of +2 = EOM of next month
    return fmtDateMalayUpper(next);
  }
  const tarikhHeader = fmtDateMalayUpper(tarikhDate);
  const dueDate = endOfNextMonth(tarikhDate); // used in point 10

  // ========= Small helpers =========
  const tw = (cm) => Math.round((cm / 2.54) * 1440); // cm → twips
  const redRun = (t, opts = {}) =>
    new TextRun({ text: t, color: "C00000", bold: true, ...opts });
  const blackRun = (t, opts = {}) => new TextRun({ text: t, ...opts });

  function makeCell(paras, opts = {}) {
    return new TableCell({
      children: Array.isArray(paras) ? paras : [paras],
      margins: { top: 100, bottom: 100, left: 0, right: 100 },
      width: { size: opts.w || 0, type: opts.wType || WidthType.AUTO },
      columnSpan: opts.span,
      shading: opts.shading,
      borders: opts.borders,
      verticalAlign: opts.vAlign,
    });
  }
  function borderAll() {
    return {
      top: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
    };
  }
  const borderNone = {
    top: { style: BorderStyle.SINGLE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.SINGLE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.SINGLE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.SINGLE, size: 0, color: "FFFFFF" },
  };

  // ========= Subject "table" with NO borders/shading =========
  const subjTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      // Title row (no shading, no borders)
      new TableRow({
        children: [
          makeCell(
            new Paragraph({
              children: [
                blackRun("PERMOHONAN KEBENARAN MERANCANG", { bold: true }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 120 }, // a little space before the detail rows
            }),
            { span: 3, borders: borderNone }
          ),
        ],
      }),

      // PROJEK : VALUE
      new TableRow({
        children: [
          makeCell(new Paragraph({ children: [blackRun("PROJEK")] }), {
            wType: WidthType.PERCENTAGE,
            w: 25,
            borders: borderNone,
          }),
          makeCell(new Paragraph({ children: [blackRun(":")] }), {
            wType: WidthType.PERCENTAGE,
            w: 5,
            borders: borderNone,
          }),
          makeCell(
            new Paragraph({
              children: [blackRun((projek || "").toString().toUpperCase())],
              spacing: { after: 60 },
            }),
            { wType: WidthType.PERCENTAGE, w: 70, borders: borderNone }
          ),
        ],
      }),

      // HARTANAH/TAPAK PROJEK : VALUE
      new TableRow({
        children: [
          makeCell(
            new Paragraph({ children: [blackRun("HARTANAH/TAPAK PROJEK")] }),
            { wType: WidthType.PERCENTAGE, w: 25, borders: borderNone }
          ),
          makeCell(new Paragraph({ children: [blackRun(":")] }), {
            wType: WidthType.PERCENTAGE,
            w: 5,
            borders: borderNone,
          }),
          makeCell(
            new Paragraph({
              children: [blackRun((tapak || "").toString().toUpperCase())],
              spacing: { after: 60 },
            }),
            { wType: WidthType.PERCENTAGE, w: 70, borders: borderNone }
          ),
        ],
      }),

      // PEMILIK/PEMAJU : VALUE
      new TableRow({
        children: [
          makeCell(new Paragraph({ children: [blackRun("PEMILIK/PEMAJU")] }), {
            wType: WidthType.PERCENTAGE,
            w: 25,
            borders: borderNone,
          }),
          makeCell(new Paragraph({ children: [blackRun(":")] }), {
            wType: WidthType.PERCENTAGE,
            w: 5,
            borders: borderNone,
          }),
          makeCell(
            new Paragraph({
              children: [blackRun((namaTetuan || "").toString().toUpperCase())],
              spacing: { after: 60 },
            }),
            { wType: WidthType.PERCENTAGE, w: 70, borders: borderNone }
          ),
        ],
      }),
    ],
  });

  // ========= Fees tables (values per screenshots) =========
  function money(n) {
    return n.toLocaleString("en-MY", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    });
  }
  function money1(n) {
    return n.toLocaleString("en-MY", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    });
  }

  const stage1 = [
    ["Landed: <=5 floors – Single Unit/Bungalow", 540],
    ["Landed: <=5 floors – Multi Unit", 1360],
    ["High Rise: >5 to <=10 floors", 2160],
    ["High Rise: >10 to <=30 floors", 2520],
    ["High Rise: >30 floors", 2520],
  ];
  const stage2 = [
    ["Landed: <=5 floors – Single Unit/Bungalow", 810],
    ["Landed: <=5 floors – Multi Unit", 2040],
    ["High Rise: >5 to <=10 floors", 2550],
    ["High Rise: >10 to <=30 floors", 2730],
    ["High Rise: >30 floors", 2960],
  ];
  const resub = [
    ["Landed: <=5 floors – Single Unit/Bungalow", 540],
    ["Landed: <=5 floors – Multi Unit", 1360],
    ["High Rise: >5 to <=10 floors", 2160],
    ["High Rise: >10 to <=30 floors", 2520],
    ["High Rise: >30 floors", 1360],
  ];

  function feeTable(title, col2Header) {
    const headerRow = new TableRow({
      children: [
        makeCell(
          new Paragraph({ children: [blackRun("No.", { bold: true })] }),
          { borders: borderAll() }
        ),
        makeCell(
          new Paragraph({
            children: [blackRun("Type of Property", { bold: true })],
          }),
          { borders: borderAll(), wType: WidthType.PERCENTAGE, w: 40 }
        ),
        makeCell(
          new Paragraph({ children: [blackRun(col2Header, { bold: true })] }),
          { borders: borderAll() }
        ),
        makeCell(
          new Paragraph({
            children: [blackRun("Amount + TAX 8%", { bold: true })],
          }),
          { borders: borderAll() }
        ),
        makeCell(
          new Paragraph({
            children: [blackRun("TOTAL INCLUSIVE TAX 8%", { bold: true })],
          }),
          { borders: borderAll() }
        ),
      ],
    });

    function rowsFrom(arr) {
      return arr.map((r, i) => {
        const amt = r[1];
        const tax = amt * 0.08;
        const total = amt * 1.08;
        return new TableRow({
          children: [
            makeCell(new Paragraph(String(i + 1)), { borders: borderAll() }),
            makeCell(new Paragraph(r[0]), {
              borders: borderAll(),
              wType: WidthType.PERCENTAGE,
              w: 40,
            }),
            makeCell(new Paragraph(money(amt)), { borders: borderAll() }),
            makeCell(new Paragraph(money1(tax)), { borders: borderAll() }),
            makeCell(new Paragraph(money1(total)), { borders: borderAll() }),
          ],
        });
      });
    }

    return [
      new Paragraph({
        children: [blackRun(title, { bold: true })],
        spacing: { before: 200, after: 100 },
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          headerRow,
          ...rowsFrom(
            title.includes("Peringkat 1")
              ? stage1
              : title.includes("Peringkat 2")
              ? stage2
              : resub
          ),
        ],
      }),
    ];
  }

  // ========= Page 1 content =========
  const page1 = [
    // Header top-left (refs/date)
    new Paragraph({
      children: [blackRun(`Ruj Kami: ${rujKami}`)],
      spacing: { after: 100 },
    }),
    new Paragraph({ children: [blackRun(`Ruj. Tuan: ${rujTuan}`)] }),
    new Paragraph({ children: [blackRun(`Tarikh: ${tarikhHeader}`)] }),

    // === NEW: Two-column header row (left = consultant + address, right = SERAHAN/FAKSIMILI) ===
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            // LEFT cell: consultant name + address (this defines where the address starts)
            makeCell(
              [
                new Paragraph({
                  children: [
                    blackRun(consultant.toString().toUpperCase(), {
                      bold: true,
                    }),
                  ],
                }),
                ...alamatConsultant
                  .toString()
                  .split(/\r?\n|,\s*/)
                  .filter(Boolean)
                  .map((ln) => new Paragraph({ children: [blackRun(ln)] })),
              ],
              {
                borders: borderNone,
                wType: WidthType.PERCENTAGE,
                w: 70, // adjust this if your address block needs more/less width
                vAlign: VerticalAlign.TOP,
              }
            ),

            // RIGHT cell: compact two-line block, right-aligned
            makeCell(
              [
                new Paragraph({
                  children: [blackRun("SERAHAN POS/TANGAN", { bold: true })],
                  alignment: AlignmentType.RIGHT,
                  spacing: { after: 0 },
                }),
                new Paragraph({
                  children: [blackRun("FAKSIMILI :", { bold: true })],
                  alignment: AlignmentType.RIGHT,
                  spacing: { after: 250 }, // keep your original bottom spacing before the next content
                }),
              ],
              {
                borders: borderNone,
                wType: WidthType.PERCENTAGE,
                w: 30,
                vAlign: VerticalAlign.TOP,
              }
            ),
          ],
        }),
      ],
    }),

    // U.P + salutation
    new Paragraph({
      children: [blackRun(`U.P: ${uP}`)],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [blackRun("Tuan/Puan,")],
      spacing: { after: 100 },
    }),

    // Subject table (with red dynamic values)
    subjTable,

    new Paragraph({
      children: [blackRun("Per:  Sokongan Merancang Pembangunan")],
      spacing: { before: 200, after: 200 }, // extra gap after border
      border: {
        bottom: {
          style: BorderStyle.SINGLE,
          size: 8,
          color: "000000",
          space: 4,
        },
      },
    }),

    // Opening paragraphs
    new Paragraph({
      children: [
        blackRun(
          "Dengan segala hormatnya saya merujuk kepada perkara tersebut di atas dan surat Arkitek/Jurutera Perunding lantikan tuan bertarikh "
        ),
        blackRun(fmtDateMalayUpper(tarikhDate), { bold: true }),
        blackRun(" berhubung perkara yang sama."),
      ],
      spacing: { after: 200 },
    }),

    // Point 2 (intro)
    new Paragraph({
      children: [
        blackRun(
          "2.  Sukacita dimaklumkan bahawa Telekom Malaysia Berhad (“TM”) bersedia untuk memberi sokongan kepada pembangunan yang dirancang bagi Projek tersebut di atas tertakluk kepada terma-terma dan syarat-syarat berikut:"
        ),
      ],
      spacing: { after: 100 },
    }),

    // (a) … (l)
    ...[
      "(a) pemilik/pemaju menyatakan tarikh mula dan jangka siap Projek tersebut;",
      "(b) pemilik/pemaju melantik perunding berdaftar untuk menyedia dan merekabentuk infrastruktur rangkaian telekomunikasi talian dari dan ke Tapak Projek;",
      "(c) mengemukakan tiga (3) set pelan susunatur yang mengandungi susunatur sivil, lurang/manholes, internal trunking dan bilik telekomunikasi/Subscriber Distribution Frame/Fiber Room (“Bilik Telekomunikasi”) berserta diagram skematik rangkaian telekomunikasi talian tetap. Pelan yang dikemukakan hendaklah mengikut spesifikasi yang telah ditetapkan oleh Suruhanjaya Komunikasi dan Multimedia Malaysia (“SKMM”) atau Malaysian Technical Standards Forum Bhd (“MTSFB”);",
      "(d) pemaju/pemilik bersetuju tidak mengenakan sebarang fi atau caj ke atas TM untuk penempatan infrastruktur telekomunikasi bagi perkhidmatan telekomunikasi yang diberikan oleh TM kepada pelanggan-pelanggannya di bangunan berkenaan sepanjang tempoh perkhidmatan tersebut disediakan;",
      "(e) menyediakan pendawaian meter elektrik berasingan bagi bekalan tenaga elektrik untuk kegunaan peralatan telekomunikasi;",
      "(f) mendapatkan kelulusan jajaran daripada Jabatan Bangunan pihak berkuasa tempatan, jika berkenaan (bagi keperluan infrastruktur untuk status MSC);",
      "(g) menyediakan pendawaian dalaman jenis gentian fiber optik single mode (bagi kawasan High Speed Broadband “HSBB” & tertakluk kepada jenis pembangunan);",
      "(h) Sila Ambil Maklum bahawa pihak TM akan mengenakan fi pemprosesan ke atas kerja-kerja semakan dan kelulusan pelan infrastruktur telekomunikasi projek yang tersebut di atas bergantung kepada jenis pembangunan yang dirancang. Fi yang dikenakan terbahagi kepada 2 peringkat seperti jadual-jadual harga di bawah:",
    ].map((t) => new Paragraph({ children: [blackRun(t)] })),

    // Peringkat 1 & 2 tables
    ...feeTable(
      "Peringkat 1: Jadual harga kerja-kerja semakan dan kelulusan pelan",
      "Stage 1 – Plan Approval Processing Fee (RM)"
    ),
    ...feeTable(
      "Peringkat 2: Jadual harga kerja-kerja ujian penerimaan (PAT) bagi tujuan pengeluaran 'Certificate of Acceptance' (COA) dan Pengesahan Kelulusan Siap Infrastruktur Telekomunikasi",
      "Stage 2 – Acceptance Test & COA Processing Fee (RM)"
    ),

    new Paragraph({
      children: [
        blackRun(
          "Nota:  Fi pemprosesan di atas tertakluk kepada Cukai Jualan dan Perkhidmatan (SST) sebanyak 8%, dan dikecualikan bagi Zon Bebas Cukai (Langkawi, Labuan dan Tioman)."
        ),
      ],
      spacing: { before: 100, after: 100 },
    }),

    new Paragraph({
      children: [
        blackRun(
          "(i) sekiranya terdapat perubahan besar (“major changes”) terhadap pelan asal Projek tersebut dan pihak tuan mengemukakan semula pelan pembangunan kepada TM untuk tujuan semakan dan kelulusan, pihak TM akan mengenakan fi pemprosesan bagi kerja-kerja semakan dan kelulusan semula seperti jadual harga berikut:"
        ),
      ],
    }),
    ...feeTable(
      "Jadual harga fi pemprosesan semula (Plan Resubmission)",
      "Plan Resubmission Processing Fee (RM)"
    ),
    new Paragraph({
      children: [
        blackRun(
          "Nota: “Major Changes” adalah merujuk kepada perubahan pelan disebabkan oleh perubahan rekabentuk pembangunan yang melibatkan semakan semula rekabentuk sivil, gentian optik dalaman dan kelulusan semula pelan telekomunikasi. Fi pemprosesan semula di atas tertakluk kepada Cukai Perkhidmatan pada kadar 8%, dan dikecualikan bagi Zon Bebas Cukai (Langkawi, Labuan dan Tioman)."
        ),
      ],
      spacing: { before: 100, after: 200 },
    }),

    ...[
      "(j) segala kerja-kerja mestilah dijalankan oleh kontraktor kerja awam yang mempunyai:",
      "      i.   Sijil sah PKK Kepala VIII – Kerja-Kerja Telekomunikasi; atau",
      "      ii.  Sijil sah CIDB di bawah kod E08 dan CE21",
      "      Pihak TM akan menyelia dan menyemak kerja-kerja semasa ujian penerimaan infrastruktur telekomunikasi dijalankan;",
      "(k) pelan yang dikemukakan mestilah mematuhi syarat-syarat seperti yang dinyatakan di dalam “Guideline On The Provision Of Basic Civil Works For Communications Infrastructure In New Development Areas (SKMM/G/01/09)” yang dikeluarkan oleh SKMM;",
      "(l) terma-terma dan syarat-syarat yang dinyatakan di dalam surat ini adalah mengikat ke atas pewaris, pengganti dalam hak milik, wasi, pentadbir semasa, wakil peribadi atau penerima serahan hak pemilik/pemaju.",
    ].map((t) => new Paragraph({ children: [blackRun(t)] })),

    ...[
      "3.  Sila Ambil Maklum bahawa sekiranya terdapat infrastruktur telekomunikasi sediada TM (“Infrastruktur Telekomunikasi TM”) di dalam kawasan projek tersebut, pihak tuan adalah bertanggungjawab sepenuhnya untuk mengalihkan Infrastruktur Telekomunikasi TM mengikut cadangan TM dan menanggung semua kos yang berkaitan dengan kerja-kerja pengalihan Infrastruktur Telekomunikasi TM terbabit.",
      "4.  Sila Ambil Maklum bahawa pihak tuan adalah bertanggungan dan perlu memastikan bahawa kerja-kerja pengalihan Infrastruktur Telekomunikasi TM dijalankan dengan teliti supaya tidak berlaku sebarang kemusnahan atau kerosakan kepada Infrastruktur Telekomunikasi TM di mana sebarang kos baik pulih beserta ganti rugi akan menjadi tanggungjawab pihak tuan.",
      "5.  Sila Ambil Maklum bahawa sekiranya semasa pihak tuan menjalankan projek tersebut di dalam dan/atau di luar kawasan pembangunan dan berlaku kemalangan yang menyebabkan kerosakan terhadap Infrastruktur Telekomunikasi TM, pihak Tuan adalah bertanggungjawab sepenuhnya untuk menanggung segala kos baik pulih berserta bayaran balik kerugian yang ditanggung oleh pihak TM.",
      "6.  Sila Ambil Maklum, bahawa semua pemasangan baru Infrastruktur Telekomunikasi TM (jika berkenaan) adalah muktamad dan tiada perancangan untuk pengalihan di masa hadapan. Namun begitu, jika terdapat sebarang keperluan dari pihak tuan untuk mengalihkan infrastruktur telekomunikasi sedia ada, TM tidak akan bertanggungjawab untuk menanggung sebarang kos pengalihan dan pemasangan semula infrastruktur tersebut. Pihak Tuan hendaklah memastikan cadangan dan perancangan Projek di atas telah mengambilkira keseluruhan perancangan pembangunan di masa hadapan.",
      "7.  Sila Ambil Maklum, pihak pemaju hendaklah memastikan bahawa setiap FTB/DP dan laluan trunking berada di bahagian luar pintu pagar kedai dan elakkan berada di kawasan terkurung.",
      "8.  Sokongan ini tidak mengikat TM untuk memberikan sebarang perkhidmatan telekomunikasi setelah kerja-kerja pembangunan siap dilaksanakan.",
      "9.  Sokongan ini hanya diperakui untuk jangkamasa 2 (dua) tahun dari tarikh surat ini. Pelan-pelan perlu dikemukakan semula untuk kelulusan selepas tamat tempoh.",
      `10. Sekiranya pihak tuan bersetuju atau tidak bersetuju dengan jadual harga, terma-terma dan syarat-syarat yang dinyatakan di atas, sila kemukakan Perakuan Persetujuan tuan (Lampiran 1) kepada kami selewat-lewatnya pada ${dueDate} untuk tindakan kami yang selanjutnya.`,
      "11. Sekiranya pihak tuan memerlukan sebarang penjelasan berkaitan dengan perkara tersebut di atas, sila hubungi Shahrin Amirul Bin Ab Rahman (shahrinamirul.abrahman@tm.com.my) di talian 011-10288084 atau Noriza Binti Ahmad Sa’don (noriza.ahmadsadon@tm.com.my) di talian 013-2068582.",
    ].map((t) => new Paragraph({ children: [blackRun(t)] })),

    new Paragraph({
      children: [blackRun("Sekian, terima kasih.")],
      spacing: { before: 200 },
    }),
    new Paragraph({
      children: [blackRun("Yang benar,")],
      spacing: { before: 100, after: 200 },
    }),
    new Paragraph({
      children: [blackRun(".............................................")],
      spacing: { after: 100 },
    }),
    new Paragraph({ children: [blackRun("AIZAN BIN ATIMAN")] }),
    new Paragraph({ children: [blackRun("Penolong Pengurus")] }),
    new Paragraph({ children: [blackRun("Telekom Malaysia Berhad")] }),
  ];

  // ========= Page 2: PERAKUAN PERSETUJUAN =========
  const persetujuanTitle = new Paragraph({
    children: [blackRun("PERAKUAN PERSETUJUAN", { bold: true })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
  });

  const headerMini = [
    new Paragraph({ children: [blackRun(`Ruj Kami: ${rujKami}`)] }),
    new Paragraph({ children: [blackRun(`Ruj. Tuan: ${rujTuan}`)] }),
    new Paragraph({
      children: [blackRun(`Tarikh: ${tarikhHeader}`)],
      spacing: { after: 100 },
    }),
  ];

  const subjTableMini = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          makeCell(
            new Paragraph({
              children: [
                blackRun("PERMOHONAN KEBENARAN MERANCANG", { bold: true }),
              ],
            }),
            { span: 3, borders: borderNone }
          ),
        ],
      }),
      new TableRow({
        children: [
          makeCell(new Paragraph("PROJEK"), { borders: borderNone }),
          makeCell(new Paragraph(": "), { borders: borderNone }),
          makeCell(
            new Paragraph({
              children: [blackRun(projek.toString().toUpperCase())],
            }),
            { borders: borderNone }
          ),
        ],
      }),
      new TableRow({
        children: [
          makeCell(new Paragraph("HARTANAH/TAPAK PROJEK"), {
            borders: borderNone,
          }),
          makeCell(new Paragraph(": "), { borders: borderNone }),
          makeCell(
            new Paragraph({
              children: [blackRun(tapak.toString().toUpperCase())],
            }),
            { borders: borderNone }
          ),
        ],
      }),
      new TableRow({
        children: [
          makeCell(new Paragraph("PEMILIK/PEMAJU"), {
            borders: {
              ...borderNone,
              bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
            },
          }),
          makeCell(new Paragraph(": "), {
            borders: {
              ...borderNone,
              bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
            },
          }),
          makeCell(
            new Paragraph({
              children: [blackRun(namaTetuan.toString().toUpperCase())],
            }),
            {
              borders: {
                ...borderNone,
                bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
              },
            }
          ),
        ],
      }),
    ],
  });

  const consentPara = new Paragraph({
    children: [
      blackRun(
        "Kami, ___________________________________ telah membaca, memahami, dan bersetuju/tidak bersetuju (potong yang tidak berkenaan) dengan jadual harga, terma-terma dan syarat-syarat yang dinyatakan di dalam surat sokongan merancang pembangunan dan lampiran berkaitan yang dikemukakan oleh pihak TM."
      ),
    ],
    spacing: { before: 200, after: 200 },
  });

  function signatureBlock(title) {
    return [
      new Paragraph({
        children: [blackRun(title, { bold: true })],
        spacing: { after: 100 },
      }),
      new Paragraph({
        children: [
          blackRun("...................................................."),
        ],
      }),
      new Paragraph({ children: [blackRun("Nama")] }),
      new Paragraph({
        children: [blackRun(": ___________________________________________")],
      }),
      new Paragraph({ children: [blackRun("Jawatan")] }),
      new Paragraph({
        children: [blackRun(": ___________________________________________")],
      }),
      new Paragraph({ children: [blackRun("Tarikh")] }),
      new Paragraph({
        children: [blackRun(": ___________________________________________")],
        spacing: { after: 200 },
      }),
    ];
  }

  // ========= Build the document =========
  const doc = new Document({
    compatibility: {
      compatibilityMode: 15, // Word 2013+
    },
    styles: {
    paragraphStyles: [
      {
        id: "Normal",
        name: "Normal",
        basedOn: "Normal",
        run: {
          font: "Calibri",  // set font
          size: 20,                 // 24 half-points = 12pt
        },
        paragraph: {
          spacing: { line: 276 },   // ~1.15 line spacing
        },
      },
    ],
  },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1"
            size: { width: 11906, height: 16838 }, // A4
          },
        },
        children: [
          ...page1,
          // ---- Page break to Lampiran ----
          new Paragraph({ children: [new TextRun({ break: 1 })] }),
          persetujuanTitle,
          ...headerMini,
          subjTableMini,
          consentPara,
          ...signatureBlock("Untuk dan Bagi pihak (Pemilik/Pemaju)"),
          ...signatureBlock("Tandatangan Saksi, (Pemilik/Pemaju)"),
        ],
      },
    ],
  });

  // ========= Download =========
  const blob = await Packer.toBlob(doc);
  const filename =
    "Surat_Penuh_" +
    (namaTetuan || "Tetuan").toString().replace(/\s+/g, "_") +
    ".docx";
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
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
