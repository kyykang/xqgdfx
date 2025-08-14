// 全局变量
let ticketData = null;
let charts = {};

// 颜色主题
const colors = {
    primary: '#667eea',
    secondary: '#764ba2',
    success: '#2ecc71',
    warning: '#f39c12',
    danger: '#e74c3c',
    info: '#3498db',
    purple: '#9b59b6',
    orange: '#e67e22',
    teal: '#1abc9c',
    pink: '#e91e63'
};

// 图表颜色数组
const chartColors = [
    colors.primary, colors.secondary, colors.success, colors.warning,
    colors.danger, colors.info, colors.purple, colors.orange,
    colors.teal, colors.pink
];

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    loadData();
    setupFileUpload();
});

// 加载数据
async function loadData() {
    try {
        const response = await fetch('ticket_data.json');
        ticketData = await response.json();
        
        // 初始化页面
        initializePage();
        initializeCharts();
        setupEventListeners();
        
    } catch (error) {
        console.error('数据加载失败:', error);
        showError('数据加载失败，请检查数据文件是否存在');
    }
}

// 初始化页面基本信息
function initializePage() {
    // 更新概览卡片
    document.getElementById('total-tickets').textContent = ticketData.summary.total_tickets.toLocaleString();
    document.getElementById('total-departments').textContent = ticketData.summary.total_departments.toLocaleString();
    document.getElementById('date-range').textContent = 
        `${ticketData.summary.date_range.start} 至 ${ticketData.summary.date_range.end}`;
    
    // 初始化全局年度筛选器
    const yearFilter = document.getElementById('year-filter');
    const years = ticketData.year_stats.labels;
    years.forEach(year => {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year + '年';
        yearFilter.appendChild(option);
    });
    
    // 初始化各个图表的年度筛选器
    initializeYearFilters();
    
    // 初始显示统计
    document.getElementById('filtered-count').textContent = 
        `显示全部 ${ticketData.summary.total_tickets.toLocaleString()} 张工单`;
}

// 初始化各个图表的年度筛选器
function initializeYearFilters() {
    const filterIds = ['dept-year-filter', 'system-year-filter', 'type-year-filter', 'status-year-filter', 'monthly-year-filter'];
    const years = ticketData.year_stats.labels;
    
    filterIds.forEach(filterId => {
        const filter = document.getElementById(filterId);
        if (filter) {
            years.forEach(year => {
                const option = document.createElement('option');
                option.value = year;
                option.textContent = year + '年';
                filter.appendChild(option);
            });
        }
    });
}

// 初始化所有图表
function initializeCharts() {
    createDepartmentChart();
    createSystemChart();
    createYearChart();
    createTypeChart();
    createStatusChart();
    createMonthlyChart();
}

// 设置事件监听器
function setupEventListeners() {
    // 全局年度筛选器
    const yearFilter = document.getElementById('year-filter');
    yearFilter.addEventListener('change', function() {
        const selectedYear = this.value;
        updateFilteredStats(selectedYear);
        updateChartsForYear(selectedYear);
    });
    
    // 各个图表的年度筛选器
    document.getElementById('dept-year-filter').addEventListener('change', function() {
        updateDepartmentChart(this.value);
    });
    
    document.getElementById('system-year-filter').addEventListener('change', function() {
        updateSystemChart(this.value);
    });
    
    document.getElementById('type-year-filter').addEventListener('change', function() {
        updateTypeChart(this.value);
    });
    
    document.getElementById('status-year-filter').addEventListener('change', function() {
        updateStatusChart(this.value);
    });
    
    document.getElementById('monthly-year-filter').addEventListener('change', function() {
        updateMonthlyChart(this.value);
    });
}

// 更新筛选统计信息
function updateFilteredStats(year) {
    const filteredCountElement = document.getElementById('filtered-count');
    
    if (year === 'all') {
        filteredCountElement.textContent = `显示全部 ${ticketData.summary.total_tickets.toLocaleString()} 张工单`;
    } else {
        const yearIndex = ticketData.year_stats.labels.indexOf(year);
        const yearCount = yearIndex >= 0 ? ticketData.year_stats.data[yearIndex] : 0;
        filteredCountElement.textContent = `${year}年共 ${yearCount.toLocaleString()} 张工单`;
    }
}

// 根据年份更新所有图表
function updateChartsForYear(year) {
    updateDepartmentChart(year);
    updateSystemChart(year);
    updateTypeChart(year);
    updateStatusChart(year);
    updateMonthlyChart(year);
    
    // 同步各个图表的筛选器
    const filterIds = ['dept-year-filter', 'system-year-filter', 'type-year-filter', 'status-year-filter', 'monthly-year-filter'];
    filterIds.forEach(filterId => {
        const filter = document.getElementById(filterId);
        if (filter) {
            filter.value = year;
        }
    });
}

// 创建部门TOP10图表
function createDepartmentChart() {
    const ctx = document.getElementById('deptChart').getContext('2d');
    const deptData = ticketData.dept_top10;
    
    charts.department = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: deptData.labels,
            datasets: [{
                label: '工单数量',
                data: deptData.data,
                backgroundColor: chartColors.slice(0, deptData.labels.length),
                borderColor: chartColors.slice(0, deptData.labels.length),
                borderWidth: 1,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `${context.parsed.y} 张工单`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(0, 0, 0, 0.1)'
                    },
                    ticks: {
                        color: '#666'
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#666',
                        maxRotation: 45
                    }
                }
            }
        }
    });
}

// 更新部门图表
function updateDepartmentChart(year) {
    if (!charts.department) return;
    
    let data;
    if (year === 'all') {
        data = ticketData.dept_top10;
    } else {
        data = ticketData.dept_by_year[year] || { labels: [], data: [] };
    }
    
    charts.department.data.labels = data.labels;
    charts.department.data.datasets[0].data = data.data;
    charts.department.data.datasets[0].backgroundColor = chartColors.slice(0, data.labels.length);
    charts.department.data.datasets[0].borderColor = chartColors.slice(0, data.labels.length);
    charts.department.update();
}

// 创建系统统计图表
function createSystemChart() {
    const ctx = document.getElementById('systemChart').getContext('2d');
    const systemData = ticketData.system_stats;
    
    charts.system = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(systemData),
            datasets: [{
                data: Object.values(systemData),
                backgroundColor: [colors.primary, colors.success, colors.warning],
                borderColor: '#fff',
                borderWidth: 3,
                hoverBorderWidth: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 20,
                        usePointStyle: true,
                        color: '#666'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed} 个 (${percentage}%)`;
                        }
                    }
                }
            },
            cutout: '60%'
        }
    });
}

// 更新系统图表
function updateSystemChart(year) {
    if (!charts.system) return;
    
    let data;
    if (year === 'all') {
        data = ticketData.system_stats;
    } else {
        data = ticketData.system_by_year[year] || { 'OA系统': 0, '营销平台': 0, 'U8C': 0 };
    }
    
    charts.system.data.labels = Object.keys(data);
    charts.system.data.datasets[0].data = Object.values(data);
    charts.system.update();
}

// 创建年度趋势图表
function createYearChart() {
    const ctx = document.getElementById('yearChart').getContext('2d');
    const yearData = ticketData.year_stats;
    
    charts.year = new Chart(ctx, {
        type: 'line',
        data: {
            labels: yearData.labels,
            datasets: [{
                label: '工单数量',
                data: yearData.data,
                borderColor: colors.primary,
                backgroundColor: colors.primary + '20',
                borderWidth: 3,
                fill: true,
                tension: 0.4,
                pointBackgroundColor: colors.primary,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointRadius: 6,
                pointHoverRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `${context.parsed.y} 张工单`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(0, 0, 0, 0.1)'
                    },
                    ticks: {
                        color: '#666'
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#666'
                    }
                }
            }
        }
    });
}

// 创建工单类型图表
function createTypeChart() {
    const ctx = document.getElementById('typeChart').getContext('2d');
    const typeData = ticketData.type_stats;
    
    charts.type = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: typeData.labels,
            datasets: [{
                data: typeData.data,
                backgroundColor: chartColors.slice(0, typeData.labels.length),
                borderColor: '#fff',
                borderWidth: 2,
                hoverBorderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 15,
                        usePointStyle: true,
                        color: '#666'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed} 个 (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// 更新工单类型图表
function updateTypeChart(year) {
    if (!charts.type) return;
    
    let data;
    if (year === 'all') {
        data = ticketData.type_stats;
    } else {
        data = ticketData.type_by_year[year] || { labels: [], data: [] };
    }
    
    charts.type.data.labels = data.labels;
    charts.type.data.datasets[0].data = data.data;
    charts.type.data.datasets[0].backgroundColor = chartColors.slice(0, data.labels.length);
    charts.type.update();
}

// 创建工单状态图表
function createStatusChart() {
    const ctx = document.getElementById('statusChart').getContext('2d');
    const statusData = ticketData.status_stats;
    
    charts.status = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: statusData.labels,
            datasets: [{
                label: '工单数量',
                data: statusData.data,
                backgroundColor: chartColors.slice(0, statusData.labels.length),
                borderColor: chartColors.slice(0, statusData.labels.length),
                borderWidth: 1,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `${context.parsed.y} 张工单`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(0, 0, 0, 0.1)'
                    },
                    ticks: {
                        color: '#666'
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#666',
                        maxRotation: 45
                    }
                }
            }
        }
    });
}

// 更新工单状态图表
function updateStatusChart(year) {
    if (!charts.status) return;
    
    let data;
    if (year === 'all') {
        data = ticketData.status_stats;
    } else {
        data = ticketData.status_by_year[year] || { labels: [], data: [] };
    }
    
    charts.status.data.labels = data.labels;
    charts.status.data.datasets[0].data = data.data;
    charts.status.data.datasets[0].backgroundColor = chartColors.slice(0, data.labels.length);
    charts.status.data.datasets[0].borderColor = chartColors.slice(0, data.labels.length);
    charts.status.update();
}

// 创建月度趋势图表
function createMonthlyChart() {
    const ctx = document.getElementById('monthlyChart').getContext('2d');
    const monthlyData = ticketData.monthly_stats;
    
    charts.monthly = new Chart(ctx, {
        type: 'line',
        data: {
            labels: monthlyData.labels,
            datasets: [{
                label: '工单数量',
                data: monthlyData.data,
                borderColor: colors.success,
                backgroundColor: colors.success + '20',
                borderWidth: 2,
                fill: true,
                tension: 0.4,
                pointBackgroundColor: colors.success,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointRadius: 4,
                pointHoverRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#2ecc71',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `${context.parsed.y} 张工单`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(0, 0, 0, 0.1)'
                    },
                    ticks: {
                        color: '#666'
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#666',
                        maxRotation: 45
                    }
                }
            }
        }
    });
}

// 更新月度趋势图表
function updateMonthlyChart(year) {
    if (!charts.monthly) return;
    
    let data;
    if (year === 'all') {
        data = ticketData.monthly_stats;
    } else {
        data = ticketData.monthly_by_year[year] || { labels: [], data: [] };
    }
    
    charts.monthly.data.labels = data.labels;
    charts.monthly.data.datasets[0].data = data.data;
    charts.monthly.update();
}

// 显示错误信息
function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.innerHTML = `
        <h3>⚠️ 错误</h3>
        <p>${message}</p>
    `;
    document.body.appendChild(errorDiv);
}

// 响应式处理
window.addEventListener('resize', function() {
    Object.values(charts).forEach(chart => {
        if (chart && typeof chart.resize === 'function') {
            chart.resize();
        }
    });
});

// 文件上传功能
function setupFileUpload() {
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const uploadBtn = document.querySelector('.upload-btn');
    const uploadStatus = document.getElementById('upload-status');
    const statusIcon = document.querySelector('.status-icon');
    const statusText = document.querySelector('.status-text');
    const progressFill = document.getElementById('progress-fill');

    // 点击上传区域触发文件选择
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 点击上传按钮触发文件选择
    uploadBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    // 文件选择处理
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            handleFileUpload(file);
        }
    });

    // 拖拽功能
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileUpload(files[0]);
        }
    });
}

// 处理文件上传
function handleFileUpload(file) {
    // 验证文件类型
    const allowedTypes = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!allowedTypes.includes(fileExtension)) {
        showUploadError('请选择Excel文件（.xlsx 或 .xls 格式）');
        return;
    }

    // 验证文件大小（限制为10MB）
    if (file.size > 10 * 1024 * 1024) {
        showUploadError('文件大小不能超过10MB');
        return;
    }

    // 显示上传状态
    showUploadProgress('正在上传文件...');

    // 创建FormData
    const formData = new FormData();
    formData.append('file', file);

    // 发送文件到后端
    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('上传失败');
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            showUploadProgress('正在处理数据...');
            // 重新加载数据
            setTimeout(() => {
                loadData().then(() => {
                    showUploadSuccess('文件上传成功，数据已更新！');
                });
            }, 1000);
        } else {
            throw new Error(data.message || '处理失败');
        }
    })
    .catch(error => {
        console.error('上传错误:', error);
        showUploadError('上传失败: ' + error.message);
    });
}

// 显示上传进度
function showUploadProgress(message) {
    const uploadStatus = document.getElementById('upload-status');
    const statusIcon = document.querySelector('.status-icon');
    const statusText = document.querySelector('.status-text');
    const progressFill = document.getElementById('progress-fill');

    uploadStatus.style.display = 'block';
    statusIcon.innerHTML = '⏳';
    statusIcon.style.animation = 'spin 1s linear infinite';
    statusText.textContent = message;
    progressFill.style.width = '60%';
}

// 显示上传成功
function showUploadSuccess(message) {
    const uploadStatus = document.getElementById('upload-status');
    const statusIcon = document.querySelector('.status-icon');
    const statusText = document.querySelector('.status-text');
    const progressFill = document.getElementById('progress-fill');

    statusIcon.innerHTML = '✅';
    statusIcon.style.animation = 'none';
    statusText.textContent = message;
    progressFill.style.width = '100%';

    // 3秒后隐藏状态
    setTimeout(() => {
        uploadStatus.style.display = 'none';
        progressFill.style.width = '0%';
    }, 3000);
}

// 显示上传错误
function showUploadError(message) {
    const uploadStatus = document.getElementById('upload-status');
    const statusIcon = document.querySelector('.status-icon');
    const statusText = document.querySelector('.status-text');
    const progressFill = document.getElementById('progress-fill');

    uploadStatus.style.display = 'block';
    statusIcon.innerHTML = '❌';
    statusIcon.style.animation = 'none';
    statusText.textContent = message;
    progressFill.style.width = '0%';

    // 5秒后隐藏状态
    setTimeout(() => {
        uploadStatus.style.display = 'none';
    }, 5000);
}