const excelURL = "https://raw.githubusercontent.com/UP3Tolis/DASHBOARD/main/NKO%20UP3%20TLI.xlsx";

let indicatorNames = [];
let monthMatrix = [];
let chart;
let filterIndicator = [];

let realizedIndicator = [];
let targetIndicator = [];
let realizedPercentageIndicator = [];

let tolisValues = [];
let tolisTrendChart;
let trendBottomRow = []; // <-- simpan nilai BW69:CH69

// Load data from Excel
async function loadExcelData() {
  try {
    const response = await fetch(excelURL);
    const arrayBuffer = await response.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    
    // Load T.UP3 sheet
    const sheet = wb.Sheets["T.UP3"];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "AT8:BL68" });

    indicatorNames = raw.map(r => r[0]);
    filterIndicator = raw.map(r => r[1]);
    monthMatrix = raw.map(r => r.slice(7, 19));

    const realizedRaw = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "X8:AP68" });
    realizedIndicator = realizedRaw.map(r => r.slice(7, 19));

    const targetRaw = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "B8:T68" });
    targetIndicator = targetRaw.map(r => r.slice(7, 19));

    targetPercentageIndicator = targetIndicator.map(row => 
      row.map(() => 100)
    );

    // baca nilai pusat (BW69:CH69) sekali dan simpan
    const bwRow = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "BW69:CH69" })[0] || [];
    trendBottomRow = bwRow.map(v => Number(v ?? 0));

    // Load trend data for multiple sheets
    const sheetNames = ["T.TOLIS", "T.BANGKIR", "T.KOTARAYA", "T.LEOK", "T.MOUTONG"];
    const trendData = {};
    const trendCharts = {};

    sheetNames.forEach(name => {
      const sh = wb.Sheets[name];
      if (!sh) {
        trendData[name] = Array(13).fill(0);
        return;
      }
      const arr = XLSX.utils.sheet_to_json(sh, { header: 1, range: "BW55:CH55" })[0] || [];
      trendData[name] = arr.map(v => Number(v ?? 0));
    });

    // Build trend charts
    sheetNames.forEach(name => buildOrUpdateTrend(name, trendData, trendCharts));

    // Initial chart update
    updateChart(0);
  } catch (error) {
    console.error("Error loading Excel data:", error);
  }
}

function buildOrUpdateTrend(name, trendData, trendCharts) {
  const canvasId = "trend-" + name;
  const el = document.getElementById(canvasId);
  if (!el) return;

  const fullLabels = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des", " "];
  const src = trendData[name] || Array(13).fill(null);
  const values = src.map(v => Number(v ?? 0));

  if (trendCharts[name]) {
    trendCharts[name].data.datasets[0].data = values;
    trendCharts[name].update();
    return;
  }

  const DPR = window.devicePixelRatio || 1;
  const visualH = parseInt(getComputedStyle(el).height, 10) || 120;
  el.style.height = visualH + 'px';
  el.width = Math.round(el.clientWidth * DPR);
  el.height = Math.round(visualH * DPR);

  trendCharts[name] = new Chart(el.getContext('2d'), {
    type: 'bar',
    data: {
      labels: fullLabels,
      datasets: [{
        data: values,
        backgroundColor: '#0077ff',
        borderColor: "#0077ff",
        borderWidth: 1,
        pointRadius: 4,
        pointBackgroundColor: "#2f00ffff",
        spanGaps: false
      }]
    },
    plugins: [ChartDataLabels],
    options: {
      responsive: false,
      maintainAspectRatio: true,
      scales: {
        y: { beginAtZero: true, max: 130, grid: { display: false }, ticks: { display: false } },
        x: { grid: { display: false }, ticks: { display: false } }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: "end",
          align: "top",
          formatter: v => v !== null ? Math.round(v) : "",
          font: { weight: "bold", size: 10 }
        }
      }
    }
  });
}

window.chartInstances = window.chartInstances || {};

// generic create/update KPI chart on given canvas id
function createKPIOnCanvas(canvasId, names, values) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  canvas.height = names.length * 18;
  const ctx = canvas.getContext("2d");

  // destroy previous instance untuk canvas ini jika ada
  if (window.chartInstances[canvasId] && typeof window.chartInstances[canvasId].destroy === "function") {
    window.chartInstances[canvasId].destroy();
  }

  const colors = values.map(v => {
    if (v >= 100) return "green";
    if (v >= 95) return "orange";
    return "red";
  });

  window.chartInstances[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: names,
      datasets: [{
        data: values,
        backgroundColor: colors,
        barPercentage: 0.5,
        categoryPercentage: 0.7
      }]
    },
    plugins: [ChartDataLabels],
    options: {
      indexAxis: 'y',
      maintainAspectRatio: false,
      layout: { padding: { right: 0, left: 0 } },
      scales: {
        x: { beginAtZero: true, max: 130, grid: { display: false }, ticks: { display: false } },
        y: { ticks: { font: { size: 11 } }, grid: { display: false } }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: "end",
          align: "right",
          formatter: value => value.toFixed(2),
          font: { size: 13, weight: "bold" },
          color: ctx => ctx.dataset.backgroundColor[ctx.dataIndex]
        }
      }
    }
  });
}

// panggil dari updateChart agar kedua canvas ter-update
function updateChart(monthIndex) {
  const filteredNames = [];
  const filteredValues = [];
  const filteredFullNames = []; // simpan nama penuh untuk selector

  let greenKPI = 0, yellowKPI = 0, redKPI = 0;
  let greenPI = 0, yellowPI = 0, redPI = 0;

  for (let i = 0; i < indicatorNames.length; i++) {
    const filterFlag = filterIndicator[i];
    const val = monthMatrix[i][monthIndex];

    if (filterFlag === "" || filterFlag === null || filterFlag === undefined) continue;

    const value = Number(val ?? 0);
    const fullLabel = indicatorNames[i] || "";                // nama penuh
    let shortLabel = fullLabel;
    if (shortLabel.length > 30) shortLabel = shortLabel.substring(0, 30) + "..."; // nama terpotong untuk chart

    filteredFullNames.push(fullLabel);
    filteredNames.push(shortLabel); // untuk chart
    filteredValues.push(value);
  }

  // simpan hasil terakhir ke global agar listener luar bisa mengakses nama penuh & nilai
  window.currentFilteredFullNames = filteredFullNames;
  window.currentFilteredValues = filteredValues;

  // Count indicators by category
  for (let i = 0; i < filteredValues.length; i++) {
    const value = filteredValues[i];

    if (i < 7) {
      if (value >= 100) greenKPI++;
      else if (value >= 95) yellowKPI++;
      else redKPI++;
    } else {
      if (value >= 100) greenPI++;
      else if (value >= 95) yellowPI++;
      else redPI++;
    }
  }

  // Update card values
  document.getElementById("greenKPIValue").innerText = greenKPI;
  document.getElementById("greenPIValue").innerText = greenPI;
  document.getElementById("yellowKPIValue").innerText = yellowKPI;
  document.getElementById("yellowPIValue").innerText = yellowPI;
  document.getElementById("redKPIValue").innerText = redKPI;
  document.getElementById("redPIValue").innerText = redPI;

  const totalIndicators = greenKPI + yellowKPI + redKPI + greenPI + yellowPI + redPI;

  // Create donut chart
  // ambil nilai tengah dari sheet jika ada (sesuai bulan)
  const centerSheetValue = (Array.isArray(trendBottomRow) && typeof trendBottomRow[monthIndex] !== 'undefined')
    ? trendBottomRow[monthIndex]
    : null;
  createDonutChart(greenKPI, yellowKPI, redKPI, greenPI, yellowPI, redPI, totalIndicators, centerSheetValue);

  // Create main bar chart (menggunakan short labels)
  createKPIOnCanvas("kpiChart", filteredNames, filteredValues);
  createKPIOnCanvas("kpiChartPage2", filteredNames, filteredValues); // untuk page 2

  // Create under-target bar chart
  createUnderTargetChart(filteredNames, filteredValues);

  // Create trend bottom chart
  createTrendBottomChart();

  // after building filteredNames and filteredValues
  // isi selector indikator (pakai nama penuh)
  const indicatorSel = document.getElementById('indicatorSelector');
  if (indicatorSel) {
    const prevValue = indicatorSel.value; // simpan pilihan sebelumnya jika ada
    indicatorSel.innerHTML = ''; // kosongkan
    filteredFullNames.forEach((name, idx) => {
      const opt = document.createElement('option');
      opt.value = String(idx); // index sama dengan nilai pada filteredValues
      opt.text = name;         // tampilkan nama penuh
      indicatorSel.appendChild(opt);
    });
    // restore previous selection jika masih valid, kalau tidak set ke 0
    if (prevValue && Array.from(indicatorSel.options).some(o => o.value === prevValue)) {
      indicatorSel.value = prevValue;
    } else {
      indicatorSel.selectedIndex = 0;
    }
    // (opsional) dispatch change agar listener lain bereaksi:
    indicatorSel.dispatchEvent(new Event('change'));
  }
}

function createDonutChart(gKPI, yKPI, rKPI, gPI, yPI, rPI, total, totalFromSheet) {
  const donutCtx = document.getElementById('donutChart').getContext('2d');
  if (window.donutChart && typeof window.donutChart.destroy === 'function') {
    window.donutChart.destroy();
  }

  // pilih nilai tengah: jika ada nilai dari sheet gunakan itu, kalau tidak pakai totalIndicators
  const centerValue = (typeof totalFromSheet === 'number' && !isNaN(totalFromSheet)) ? totalFromSheet : total;
  const totalFormatted = Number(centerValue).toLocaleString('id-ID', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

  window.donutChart = new Chart(donutCtx, {
    type: 'doughnut',
    data: {
      labels: ['Green KPI', 'Yellow KPI', 'Red KPI'],
      datasets: [{
        data: [gKPI + gPI, yKPI + yPI, rKPI + rPI],
        backgroundColor: ['#4caf50', '#ff9800', '#f44336'],
        hoverOffset: 4
      }]
    },
    plugins: [
      {
        maintainAspectRatio: true,
        id: 'centerText',
        beforeDraw(chart, args, options) {
          const { ctx, chartArea: { left, right, top, bottom } } = chart;
          const centerX = (left + right) / 2;
          const centerY = (top + bottom) / 2;
          ctx.save();

          ctx.fillStyle = options.color || '#333';
          ctx.font = `${options.fontWeight || 'bold'} ${options.fontSize || '40px'} ${options.fontFamily || 'Arial'}`;
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';
          ctx.fillText(options.text || '', centerX, centerY - (options.offsetY || 6));

          if (options.subtext) {
            ctx.fillStyle = options.subColor || '#0077cc';
            ctx.font = `${options.subFontWeight || '600'} ${options.subFontSize || '14px'} ${options.fontFamily || 'Arial'}`;
            ctx.fillText(options.subtext, centerX, centerY + (options.subOffsetY || 26));
          }

          ctx.restore();
        }
      }
    ],
    options: {
      responsive: false,
      cutout: '70%',
      elements: { arc: { borderWidth: 2 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function (tooltipItem) {
              const label = tooltipItem.label || '';
              const value = tooltipItem.raw || 0;
              return `${label}: ${value} (${((value / total) * 100).toFixed(2)}%)`;
            }
          }
        },
        datalabels: { display: false },
        centerText: {
          text: totalFormatted,
          subtext: '',
          color: '#333',
          subColor: '#0077cc',
          fontSize: '40px',
          subFontSize: '14px'
        }
      }
    }
  });
}

function createUnderTargetChart(names, values) {
  const underTargetNames = [];
  const underTargetValues = [];
  const underTargetColors = [];

  names.forEach((name, i) => {
    if (values[i] < 100) {
      underTargetNames.push(name);
      underTargetValues.push(values[i]);
      underTargetColors.push(values[i] >= 95 ? "orange" : "red");
    }
  });

  const canvas = document.getElementById('underTargetChart');
  if (!canvas) return;

  canvas.height = underTargetNames.length * 18;

  const ctx = canvas.getContext('2d');
  if (window.underTargetChart && typeof window.underTargetChart.destroy === 'function') {
    window.underTargetChart.destroy();
  }

  window.underTargetChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: underTargetNames,
      datasets: [{
        data: underTargetValues,
        backgroundColor: underTargetColors,
        barPercentage: 0.7,
        categoryPercentage: 0.7
      }]
    },
    plugins: [ChartDataLabels],
    options: {
      indexAxis: 'y',
      maintainAspectRatio: false,
      layout: { padding: { right: 0, left: 0 } },
      scales: {
        x: { beginAtZero: true, max: 125, grid: { display: false }, ticks: { display: false } },
        y: { ticks: { font: { size: 11 } }, grid: { display: false } }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: 'end',
          align: 'right',
          formatter: value => value.toFixed(2),
          font: { size: 11, weight: 'bold' },
          color: ctx => ctx.dataset.backgroundColor[ctx.dataIndex]
        }
      }
    }
  });
}

async function createTrendBottomChart() {
  try {
    const response = await fetch(excelURL);
    const arrayBuffer = await response.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = wb.Sheets["T.UP3"];

    const trendBottomData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "BW69:CH69" })[0] || [];
    const monthLabels = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    const trendValues = trendBottomData.map(v => Number(v ?? 0));

    const canvas = document.getElementById('trendBottomChart');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    if (window.trendBottomChart && typeof window.trendBottomChart.destroy === 'function') {
      window.trendBottomChart.destroy();
    }

    window.trendBottomChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: monthLabels,
        datasets: [{
          label: 'Tren Indikator',
          data: trendValues,
          backgroundColor: '#0077ff',
          borderColor: '#0055cc',
          borderWidth: 1,
          barPercentage: 0.5,
          categoryPercentage: 0.7
        }]
      },
      plugins: [ChartDataLabels],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: { beginAtZero: true, max: 130, grid: { display: false }, ticks: { display: false } },
          x: { grid: { display: false }, ticks: { font: { size: 11 } } }
        },
        plugins: {
          legend: { display: false },
          datalabels: {
            anchor: 'end',
            align: 'top',
            formatter: value => value.toFixed(2),
            font: { size: 10, weight: 'bold' },
            color: '#000'
          }
        }
      }
    });
  } catch (error) {
    console.error("Error creating trend bottom chart:", error);
  }
}

function indicatorChart(indicatorSel, indicatorNames, percentage, values, targetValue, canvasId) {
  const chart_realizedIndicator = [];
  const chart_targetIndicator = [];
  const chart_percentageIndicator = [];

  for (let i = 0; i < indicatorNames.length; i++) {
    if (indicatorSel == indicatorNames[i]) {
      chart_realizedIndicator.push(values[i]);
      chart_targetIndicator.push(targetValue[i]);
      chart_percentageIndicator.push(percentage[i]);

      console.log("indicator matched:", indicatorSel, "==", chart_realizedIndicator[0]);
    }
  }

  maxValue = 0;
  const maxValueRealized = Math.max(...chart_realizedIndicator[0]);
  const maxValueTarget = Math.max(...chart_targetIndicator[0]);

  if (maxValueRealized > maxValueTarget) {
    maxValue = maxValueRealized;
  } else {
    maxValue = maxValueTarget;
  }

  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  if (!window.indicatorChartInstances) {
    window.indicatorChartInstances = {};
  }

  if (window.indicatorChartInstances[canvasId]) {
    window.indicatorChartInstances[canvasId].destroy();
  }

  const colors = chart_percentageIndicator[0].map(v => {
    if (v >= 100) return "green";
    if (v >= 95) return "orange";
    return "red";
  });

  window.indicatorChartInstances[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"],
      datasets: [
      {
        label: "Realisasi Indikator",
        data: chart_realizedIndicator[0],
        borderColor: "#0077ff",
        backgroundColor: colors,
        borderWidth: 0,
        barPercentage: 0.5,
        categoryPercentage: 0.7,
        order: 2
      },
      {
        // Line chart - Target (opsional, bisa disesuaikan)
          type: "line",
          label: "Target Indikator",
          data: chart_targetIndicator[0],  // atau sesuai target aktual
          borderColor: "#f44336",
          backgroundColor: "rgba(244, 67, 54, 0.1)",
          borderWidth: 2,
          pointRadius: 5,
          pointBackgroundColor: "#f44336",
          pointBorderColor: "#fff",
          pointBorderWidth: 2,
          tension: 0.4,  // smoothness kurva
          order: 1,  // tampilkan di depan bar
          datalabels: { display: false}  // sembunyikan data label untuk line
      }]
    },
    plugins: [ChartDataLabels],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: { beginAtZero: true, max: maxValue + 10, grid: { display: false }, ticks: { display: false } },
        x: { grid: { display: false }, ticks: { font: { size: 11 } } }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: 'end',
          align: 'top',
          formatter: value => value.toFixed(2),
          font: { size: 12, weight: 'bold' },
          color: '#000'
        },
      }
    }
  });
        
    
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
  console.log("DOM Loaded - Starting initialization");
  
  loadExcelData();
  showPage('dashboard', null);

  // attach listeners to all month selectors (sync them)
  document.querySelectorAll('.monthSelector').forEach(sel => {
    sel.addEventListener('change', function(e) {
      const m = parseInt(e.target.value);
      // sync value to all selectors
      document.querySelectorAll('.monthSelector').forEach(s => { if (s.value !== String(m)) s.value = String(m); });
      console.log("Month changed to:", m);
      updateChart(m);
    });
  });

  const indicatorSel = document.getElementById('indicatorSelector');
  if (indicatorSel) {
    indicatorSel.addEventListener('change', function(e) {
      const idx = parseInt(e.target.value);
      const names = window.currentFilteredFullNames || [];
      const vals = window.currentFilteredValues || [];
      console.log('Indicator selected:', names[idx], vals[idx]);

      indicatorChart(names[idx], indicatorNames, monthMatrix, monthMatrix, targetPercentageIndicator, "indicatorChartCanvas");
      indicatorChart(names[idx], indicatorNames, monthMatrix, realizedIndicator, targetIndicator, "indicator2ChartCanvas");
    });
  }
});

// Navigation function
function showPage(pageId, clickElement) {
  const pages = document.querySelectorAll('.page');
  pages.forEach(page => page.classList.remove('active'));

  const activePage = document.getElementById(pageId);
  if (activePage) {
    activePage.classList.add('active');
  }

  const links = document.querySelectorAll('ul li a');
  links.forEach(link => link.classList.remove('active'));
  
  // Tambahkan active class ke link yang diklik (jika ada)
  if (clickElement) {
    clickElement.classList.add('active');
  }
}