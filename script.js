// ==========================================
// BAHAGIAN 1 - INISIALISASI
// ==========================================

// Ambil data murid dari Local Storage.
// Jika tiada, gunakan data asal dalam students.js
let studentList = JSON.parse(localStorage.getItem("students")) || students;

// Elemen HTML
const tbody = document.querySelector("#studentTable tbody");
const searchInput = document.querySelector("#searchInput");
const assessmentBody = document.querySelector("#assessmentTable tbody");


// ==========================================
// PAPAR SENARAI MURID
// ==========================================

function displayStudents(list) {

    tbody.innerHTML = "";

    list.forEach(student => {

        tbody.innerHTML += `
        <tr>

            <td>${student.id}</td>

            <td>${student.name}</td>

            <td>${student.gender}</td>

            <td>
                <button>Rekod</button>
            </td>

        </tr>
        `;

    });

}


// ==========================================
// PAPAR JADUAL PENTAKSIRAN
// ==========================================

function displayAssessment(list) {

    assessmentBody.innerHTML = "";

    list.forEach(student => {

        assessmentBody.innerHTML += `
        <tr>

            <td>${student.id}</td>

            <td>${student.name}</td>

            <td>

                <select class="tp">

                    <option value="">-</option>
                    <option>TP1</option>
                    <option>TP2</option>
                    <option>TP3</option>
                    <option>TP4</option>
                    <option>TP5</option>
                    <option>TP6</option>

                </select>

            </td>

        </tr>
        `;

    });

}


// ==========================================
// CARIAN MURID
// ==========================================

searchInput.addEventListener("keyup", function () {

    const keyword = this.value.toLowerCase();

    const filtered = studentList.filter(student =>
        student.name.toLowerCase().includes(keyword)
    );

    displayStudents(filtered);

});


// ==========================================
// PAPAR DATA KALI PERTAMA
// ==========================================

displayStudents(studentList);
displayAssessment(studentList);


// ==========================================
// UNIT SEMASA & STANDARD PEMBELAJARAN
// ==========================================

const unitSelect = document.getElementById("unit");
const currentUnit = document.getElementById("currentUnit");
const spSelect = document.getElementById("sp");

// Papar unit semasa apabila aplikasi dibuka
currentUnit.textContent = unitSelect.value;

function updateSP() {

    spSelect.innerHTML = "";

    const list = spData[unitSelect.value];

    if (!list) return;

    list.forEach(sp => {

        const option = document.createElement("option");

        option.value = sp;
        option.textContent = sp;

        spSelect.appendChild(option);

    });

}

// Papar SP kali pertama
updateSP();

// Bila Unit berubah
unitSelect.addEventListener("change", () => {

    currentUnit.textContent = unitSelect.value;

    updateSP();

});

// ==========================================
// BAHAGIAN 2 - TETAPKAN SEMUA TP
// ==========================================

function setAllTP(tpValue) {

    document.querySelectorAll(".tp").forEach(tp => {
        tp.value = tpValue;
    });

}

document.getElementById("tp1All").addEventListener("click", () => setAllTP("TP1"));
document.getElementById("tp2All").addEventListener("click", () => setAllTP("TP2"));
document.getElementById("tp3All").addEventListener("click", () => setAllTP("TP3"));
document.getElementById("tp4All").addEventListener("click", () => setAllTP("TP4"));
document.getElementById("tp5All").addEventListener("click", () => setAllTP("TP5"));
document.getElementById("tp6All").addEventListener("click", () => setAllTP("TP6"));

document.getElementById("clearAll").addEventListener("click", () => setAllTP(""));



// ==========================================
// BAHAGIAN 3 - SIMPAN REKOD PENTAKSIRAN
// ==========================================

document.getElementById("saveAssessment").addEventListener("click", () => {

    // Cipta rekod baru
    const record = {

        unit: document.getElementById("unit").value,

        sp: document.getElementById("sp").value,

        date: new Date().toLocaleDateString("ms-MY"),

        students: []

    };

    // Ambil semua TP murid
    document.querySelectorAll("#assessmentTable tbody tr").forEach((row, index) => {

        record.students.push({

            id: studentList[index].id,

            name: studentList[index].name,

            gender: studentList[index].gender,

            tp: row.querySelector(".tp").value

        });

    });

    // Ambil rekod lama
    let assessments = JSON.parse(localStorage.getItem("assessments")) || [];

    // Tambah rekod baru
    assessments.push(record);

    // Simpan semula
    localStorage.setItem("assessments", JSON.stringify(assessments));

    alert("✅ Rekod berjaya disimpan!");

});

console.log(JSON.parse(localStorage.getItem("assessments")));

// ==========================================
// PAPAR DATA LOCAL STORAGE
// ==========================================

document.getElementById("showData").addEventListener("click", () => {

    const data = JSON.parse(localStorage.getItem("assessments")) || [];

    document.getElementById("output").textContent =
        JSON.stringify(data, null, 2);

});

// ==========================================
// BAHAGIAN 4 - PAPAR REKOD
// ==========================================

document.getElementById("loadRecords").addEventListener("click", () => {

    const tbody = document.querySelector("#recordTable tbody");

    tbody.innerHTML = "";

    let records = JSON.parse(localStorage.getItem("assessments")) || [];

    records.forEach(record => {

        tbody.innerHTML += `
        <tr>

            <td>${record.date}</td>

            <td>${record.unit}</td>

            <td>${record.sp}</td>

            <td>${record.students.length}</td>

        </tr>
        `;

    });

});

// ==========================================
// BAHAGIAN 5 - IMPORT EXCEL
// ==========================================

document.getElementById("importBtn").addEventListener("click", () => {

    const fileInput = document.getElementById("excelFile");

    if (fileInput.files.length === 0) {
        alert("Sila pilih fail Excel dahulu.");
        return;
    }

    const file = fileInput.files[0];

    const reader = new FileReader();

    reader.onload = function (e) {

        const data = new Uint8Array(e.target.result);

        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const excelData = XLSX.utils.sheet_to_json(sheet);

studentList = excelData.map(row => ({

    id: row["ID"],

    name: (row["NAMA"] || row["NAMA "] || "").trim(),

    gender: row["JANTINA"]

}));

        // Simpan ke Local Storage
        localStorage.setItem("students", JSON.stringify(studentList));

        // Papar semula jadual
        displayStudents(studentList);
        displayAssessment(studentList);

        // Kemaskini bilangan murid
        document.getElementById("studentCount").textContent =
            studentList.length + " Orang";

        alert("✅ Import Excel berjaya!");

    };

    reader.readAsArrayBuffer(file);

});

// ==========================================
// BAHAGIAN 6 - RESET DATA MURID
// ==========================================

document.getElementById("resetBtn").addEventListener("click", () => {

    localStorage.removeItem("students");

    studentList = students;

    displayStudents(studentList);
    displayAssessment(studentList);

    document.getElementById("studentCount").textContent =
        studentList.length + " Orang";

    alert("✅ Data murid telah direset.");

});