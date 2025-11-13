const excelURL = "https://raw.githubusercontent.com/UP3Tolis/DASHBOARD/main/NKO%20UP3%20TLI.xlsx";

let indicatorNames = [];
let monthMatrix = [];
let chart;
let filterIndicator = [];

let tolisValues = [];
let tolisTrendChart;

// Load data from Excel
async function loadExcelData() {
  try {
    const response = await fetch(excelURL);
    const arrayBuffer = await response.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    
    // Load T.UP3 sheet
    const sheet = wb.Sheets["T.UP3"];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "AT8:BL62" });

    indicatorNames = raw.map(r => r[0]);
    filterIndicator = raw.map(r => r[1]);
    monthMatrix = raw.map(r => r.slice(7, 19));

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
    type: 'line',
    data: {
      labels: fullLabels,
      datasets: [{
        data: values,
        borderColor: "#0077ff",
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: "#0077ff",
        spanGaps: false
      }]
    },
    plugins: [ChartDataLabels],
    options: {
      responsive: false,
      maintainAspectRatio: true,
      scales: {
        y: { beginAtZero: true, max: 120, grid: { display: false }, ticks: { display: false } },
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

function updateChart(monthIndex) {
  const filteredNames = [];
  const filteredValues = [];

  let greenKPI = 0, yellowKPI = 0, redKPI = 0;
  let greenPI = 0, yellowPI = 0, redPI = 0;

  for (let i = 0; i < indicatorNames.length; i++) {
    const filterFlag = filterIndicator[i];
    const val = monthMatrix[i][monthIndex];

    if (filterFlag === "" || filterFlag === null || filterFlag === undefined) continue;

    const value = Number(val ?? 0);
    let label = indicatorNames[i];
    if (label.length > 30) label = label.substring(0, 30) + "...";

    filteredNames.push(label);
    filteredValues.push(value);
  }

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
  createDonutChart(greenKPI, yellowKPI, redKPI, greenPI, yellowPI, redPI, totalIndicators);

  // Create main bar chart
  createKPIChart(filteredNames, filteredValues);

  // Create under-target bar chart
  createUnderTargetChart(filteredNames, filteredValues);

  // Create trend bottom chart
  createTrendBottomChart();
}

function createDonutChart(gKPI, yKPI, rKPI, gPI, yPI, rPI, total) {
  const donutCtx = document.getElementById('donutChart').getContext('2d');
  if (window.donutChart && typeof window.donutChart.destroy === 'function') {
    window.donutChart.destroy();
  }

  const totalFormatted = total.toLocaleString('id-ID', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

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
          subtext: 'INDICATOR',
          color: '#333',
          subColor: '#0077cc',
          fontSize: '40px',
          subFontSize: '14px'
        }
      }
    }
  });
}

function createKPIChart(names, values) {
  const colors = values.map(v => {
    if (v >= 100) return "green";
    if (v >= 95) return "orange";
    return "red";
  });

  const canvas = document.getElementById("kpiChart");
  canvas.height = names.length * 18;

  if (chart) chart.destroy();
  const ctx = canvas.getContext("2d");

  chart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: names,
      datasets: [{
        data: values,
        backgroundColor: colors,
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
        x: { beginAtZero: true, max: 130, grid: { display: false }, ticks: { display: false } },
        y: { ticks: { font: { size: 11 } }, grid: { display: false } }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: "end",
          align: "right",
          formatter: value => value.toFixed(2),
          font: { size: 11, weight: "bold" },
          color: ctx => ctx.dataset.backgroundColor[ctx.dataIndex]
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

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
  console.log("DOM Loaded - Starting initialization");
  
  loadExcelData();
  showPage('dashboard', null);  // Kirim null untuk initial load

  // Tambahkan event listener untuk monthSelector
  const monthSelector = document.getElementById("monthSelector");
  if (monthSelector) {
    monthSelector.addEventListener("change", function(e) {
      const m = parseInt(e.target.value);
      console.log("Month changed to:", m);
      updateChart(m);
    });
  } else {
    console.error("monthSelector element not found!");
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