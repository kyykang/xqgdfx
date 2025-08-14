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
    document.getElementById('total-departments').textContent = ticketData.summary.total_departments;
    document.getElementById('date-range').textContent = `${ticketData.summary.date_range.start} 至 ${ticketData.summary.date_range.end}`;
    
    // 初始化年份筛选器
    const yearFilter = document.getElementById('year-filter');
    const years = ticketData.year_stats.labels.slice().reverse();
    
    years.forEach(year => {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = `${year}年`;
        yearFilter.appendChild(option);
    });
    
    // 更新时间
    document.getElementById('update-time').textContent = new Date().toLocaleString('zh-CN');
    
    // 初始显示统计
    updateFilteredStats('all');
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
    const yearFilter = document.getElementById('year-filter');
    yearFilter.addEventListener('change', function() {
        const selectedYear = this.value;
        updateFilteredStats(selectedYear);
        updateChartsForYear(selectedYear);
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

// 根据年份更新图表
function updateChartsForYear(year) {
    // 这里可以根据需要实现年份筛选功能
    // 目前保持原有图表显示
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
                borderWidth: 2,
                borderRadius: 8,
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
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `工单数量: ${context.parsed.y.toLocaleString()}`;
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

// 创建系统使用统计图表
function createSystemChart() {
    const ctx = document.getElementById('systemChart').getContext('2d');
    const systemData = ticketData.system_stats;
    
    charts.system = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['OA系统', '营销平台', 'U8C'],
            datasets: [{
                data: [systemData.OA系统, systemData.营销平台, systemData.U8C],
                backgroundColor: [colors.primary, colors.success, colors.warning],
                borderColor: ['white', 'white', 'white'],
                borderWidth: 3,
                hoverOffset: 10
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
                        font: {
                            size: 12
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed.toLocaleString()} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// 创建年度趋势图表
function createYearChart() {
    const ctx = document.getElementById('yearChart').getContext('2d');
    const yearlyData = ticketData.year_stats;
    
    charts.year = new Chart(ctx, {
        type: 'line',
        data: {
            labels: yearlyData.labels.map(year => `${year}年`),
            datasets: [{
                label: '工单数量',
                data: yearlyData.data,
                borderColor: colors.primary,
                backgroundColor: colors.primary + '20',
                borderWidth: 3,
                fill: true,
                tension: 0.4,
                pointBackgroundColor: colors.primary,
                pointBorderColor: 'white',
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
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `工单数量: ${context.parsed.y.toLocaleString()}`;
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

// 创建工单类型分布图表
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
                borderColor: 'white',
                borderWidth: 2,
                hoverOffset: 8
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
                        font: {
                            size: 11
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed.toLocaleString()} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// 创建工单状态统计图表
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
                backgroundColor: [colors.success, colors.warning, colors.danger, colors.info],
                borderColor: [colors.success, colors.warning, colors.danger, colors.info],
                borderWidth: 2,
                borderRadius: 8,
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
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `工单数量: ${context.parsed.y.toLocaleString()}`;
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

// 创建月度趋势图表
function createMonthlyChart() {
    const ctx = document.getElementById('monthlyChart').getContext('2d');
    const monthlyData = ticketData.monthly_stats;
    
    charts.monthly = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: monthlyData.labels.map(month => {
                const [year, monthNum] = month.split('-');
                return `${year}年${monthNum}月`;
            }),
            datasets: [{
                label: '工单数量',
                data: monthlyData.data,
                backgroundColor: colors.secondary + '80',
                borderColor: colors.secondary,
                borderWidth: 2,
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
                    titleColor: 'white',
                    bodyColor: 'white',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            return `工单数量: ${context.parsed.y.toLocaleString()}`;
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
                        maxRotation: 45,
                        callback: function(value, index) {
                            // 只显示部分标签以避免拥挤
                            return index % 3 === 0 ? this.getLabelForValue(value) : '';
                        }
                    }
                }
            }
        }
    });
}

// 显示错误信息
function showError(message) {
    const container = document.querySelector('.container');
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.innerHTML = `
        <div style="background: #e74c3c; color: white; padding: 20px; border-radius: 10px; text-align: center; margin: 20px 0;">
            <h3>⚠️ 错误</h3>
            <p>${message}</p>
        </div>
    `;
    container.appendChild(errorDiv);
}

// 窗口大小改变时重新调整图表
window.addEventListener('resize', function() {
    Object.values(charts).forEach(chart => {
        if (chart) {
            chart.resize();
        }
    });
});