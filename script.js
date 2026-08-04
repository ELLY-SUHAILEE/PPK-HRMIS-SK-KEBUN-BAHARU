let allStaffData = [];
let guruChartInstance = null;
let akpChartInstance = null;

document.getElementById('excelFile').addEventListener('change', handleFile, false);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Baca data dari baris ke-6 (Header bermula di Row 5/6 Excel)
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 4 });

        processHRMISData(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

function processHRMISData(data) {
    allStaffData = [];

    let guruTotal = 0, guruAchieved = 0, guruNotAchieved = 0;
    let akpTotal = 0, akpAchieved = 0, akpNotAchieved = 0;

    data.forEach(row => {
        // Ambil data dari lajur spesifik Excel anda
        const rawNameKp = row['NO. KP / NAMA'] || row['NAMA'] || '';
        const categoryRaw = (row['KATEGORI'] || '').toString().trim().toUpperCase();
        
        // Dapatkan Nilai Jam
        let hours = 0;
        if (row['JUMLAH JAM_1'] !== undefined) {
            hours = parseFloat(row['JUMLAH JAM_1']);
        } else if (row['JUMLAH JAM'] !== undefined) {
            if (typeof row['JUMLAH JAM'] === 'number') {
                hours = row['JUMLAH JAM'];
            } else {
                // Asingkan nombor daripada teks "178 jam 54 minit"
                const match = row['JUMLAH JAM'].toString().match(/(\d+)\s*jam/i);
                hours = match ? parseFloat(match[1]) : parseFloat(row['JUMLAH JAM']) || 0;
            }
        }

        // Jika tiada nama atau data kosong, abaikan
        if (!rawNameKp || rawNameKp.toString().includes('LAPORAN')) return;

        // Asingkan No. KP dan Nama daripada "921022035658 /\nCHE SITI NOR..."
        let name = rawNameKp;
        let ic = '-';
        if (rawNameKp.includes('/')) {
            const parts = rawNameKp.split('/');
            ic = parts[0].trim();
            name = parts[1].replace(/\n/g, ' ').trim();
        }

        const category = categoryRaw.includes('AKP') ? 'AKP' : 'GURU';
        const isAchieved = hours >= 40;

        if (category === 'GURU') {
            guruTotal++;
            if (isAchieved) guruAchieved++;
            else guruNotAchieved++;
        } else {
            akpTotal++;
            if (isAchieved) akpAchieved++;
            else akpNotAchieved++;
        }

        allStaffData.push({
            ic: ic,
            name: name,
            category: category,
            hours: isNaN(hours) ? 0 : hours,
            isAchieved: isAchieved
        });
    });

    // Pengiraan Peratusan
    const guruAchievedPct = guruTotal > 0 ? ((guruAchieved / guruTotal) * 100).toFixed(1) : 0;
    const guruNotAchievedPct = guruTotal > 0 ? ((guruNotAchieved / guruTotal) * 100).toFixed(1) : 0;

    const akpAchievedPct = akpTotal > 0 ? ((akpAchieved / akpTotal) * 100).toFixed(1) : 0;
    const akpNotAchievedPct = akpTotal > 0 ? ((akpNotAchieved / akpTotal) * 100).toFixed(1) : 0;

    // Kemaskini Kad GURU
    document.getElementById('guruTotal').innerText = guruTotal;
    document.getElementById('guruAchieved').innerText = guruAchieved;
    document.getElementById('guruAchievedPct').innerText = `${guruAchievedPct}% mencapai target`;
    document.getElementById('guruNotAchieved').innerText = guruNotAchieved;
    document.getElementById('guruNotAchievedPct').innerText = `${guruNotAchievedPct}% belum mencapai`;

    // Kemaskini Kad AKP
    document.getElementById('akpTotal').innerText = akpTotal;
    document.getElementById('akpAchieved').innerText = akpAchieved;
    document.getElementById('akpAchievedPct').innerText = `${akpAchievedPct}% mencapai target`;
    document.getElementById('akpNotAchieved').innerText = akpNotAchieved;
    document.getElementById('akpNotAchievedPct').innerText = `${akpNotAchievedPct}% belum mencapai`;

    // Render Jadual & Carta
    renderTable(allStaffData);
    renderCharts(guruAchieved, guruNotAchieved, akpAchieved, akpNotAchieved);
}

function renderTable(dataList) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (dataList.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">Tiada data ditemui.</td></tr>';
        return;
    }

    dataList.forEach(item => {
        const tr = document.createElement('tr');
        const katBadge = item.category === 'GURU' 
            ? '<span class="badge badge-guru">GURU</span>' 
            : '<span class="badge badge-akp">AKP</span>';
        
        const statusBadge = item.isAchieved 
            ? '<span class="badge badge-success">Mencapai (≥ 40 Jam)</span>' 
            : '<span class="badge badge-danger">Belum Capai (< 40 Jam)</span>';

        tr.innerHTML = `
            <td style="font-weight: 600;">${item.name}</td>
            <td>${item.ic}</td>
            <td>${katBadge}</td>
            <td style="font-weight: bold; color: ${item.isAchieved ? '#16a34a' : '#dc2626'};">${item.hours} jam</td>
            <td>${statusBadge}</td>
        `;
        tbody.appendChild(tr);
    });
}

function filterTable() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    const filtered = allStaffData.filter(item => 
        item.name.toLowerCase().includes(query) || item.ic.toLowerCase().includes(query)
    );
    renderTable(filtered);
}

function renderCharts(guruAchieved, guruNotAchieved, akpAchieved, akpNotAchieved) {
    // Carta Guru
    const ctxGuru = document.getElementById('guruChart').getContext('2d');
    if (guruChartInstance) guruChartInstance.destroy();

    guruChartInstance = new Chart(ctxGuru, {
        type: 'pie',
        data: {
            labels: ['Mencapai ≥ 40 Jam', 'Belum Mencapai < 40 Jam'],
            datasets: [{
                data: [guruAchieved, guruNotAchieved],
                backgroundColor: ['#22c55e', '#ef4444'],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let total = context.dataset.data.reduce((a, b) => a + b, 0);
                            let value = context.raw;
                            let percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                            return ` ${context.label}: ${value} orang (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });

    // Carta AKP
    const ctxAkp = document.getElementById('akpChart').getContext('2d');
    if (akpChartInstance) akpChartInstance.destroy();

    akpChartInstance = new Chart(ctxAkp, {
        type: 'pie',
        data: {
            labels: ['Mencapai ≥ 40 Jam', 'Belum Mencapai < 40 Jam'],
            datasets: [{
                data: [akpAchieved, akpNotAchieved],
                backgroundColor: ['#10b981', '#f43f5e'],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let total = context.dataset.data.reduce((a, b) => a + b, 0);
                            let value = context.raw;
                            let percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                            return ` ${context.label}: ${value} orang (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}
