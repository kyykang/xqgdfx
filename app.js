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
    
    // 计算并显示草稿数量
    const totalTickets = ticketData.summary.total_tickets;
    const draftCount = ticketData.audit_stats.data[ticketData.audit_stats.labels.indexOf('草稿')];
    document.getElementById('draft-count').textContent = `其中草稿: ${draftCount.toLocaleString()}`;
    
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
    
    // 由于默认选中排除草稿，需要更新所有图表显示排除草稿的数据
    updateDepartmentChart('all');
    updateSystemChart('all');
    updateYearChart();
    updateTypeChart('all');
    updateStatusChart('all');
    updateMonthlyChart('all');
}

// 设置事件监听器
function setupEventListeners() {
    // 全局年度筛选器
    const yearFilter = document.getElementById('year-filter');
    const halfYearFilter = document.getElementById('half-year-filter');
    const halfYearLabel = document.getElementById('half-year-label');
    
    yearFilter.addEventListener('change', function() {
        const selectedYear = this.value;
        
        // 显示或隐藏半年度筛选器
        if (selectedYear === 'all') {
            halfYearFilter.style.display = 'none';
            halfYearLabel.style.display = 'none';
            halfYearFilter.value = 'all'; // 重置为全年
        } else {
            halfYearFilter.style.display = 'inline-block';
            halfYearLabel.style.display = 'inline-block';
        }
        
        updateFilteredStats(selectedYear, halfYearFilter.value);
        updateChartsForYear(selectedYear);
    });
    
    // 半年度筛选器
    halfYearFilter.addEventListener('change', function() {
        const selectedYear = yearFilter.value;
        const selectedHalfYear = this.value;
        updateFilteredStats(selectedYear, selectedHalfYear);
        updateChartsForYear(selectedYear);
    });
    
    // 各个图表的年度筛选器
    const deptYearFilter = document.getElementById('dept-year-filter');
    if (deptYearFilter) {
        deptYearFilter.addEventListener('change', function() {
            const selectedYear = this.value;
            const halfYearFilter = document.getElementById('dept-half-year-filter');
            const halfYearLabel = document.getElementById('dept-half-year-label');
            
            // 显示或隐藏半年度筛选器
            if (halfYearFilter && halfYearLabel) {
                if (selectedYear === 'all') {
                    halfYearFilter.style.display = 'none';
                    halfYearLabel.style.display = 'none';
                    halfYearFilter.value = 'all';
                } else {
                    halfYearFilter.style.display = 'inline-block';
                    halfYearLabel.style.display = 'inline-block';
                }
            }
            
            const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
            updateDepartmentChart(selectedYear, selectedHalfYear);
        });
    }
    
    const deptHalfYearFilter = document.getElementById('dept-half-year-filter');
    if (deptHalfYearFilter) {
        deptHalfYearFilter.addEventListener('change', function() {
            const selectedYear = document.getElementById('dept-year-filter').value;
            const selectedHalfYear = this.value;
            console.log('部门图表半年度筛选器变化:', selectedYear, selectedHalfYear);
            updateDepartmentChart(selectedYear, selectedHalfYear);
        });
    }
    
    const systemYearFilter = document.getElementById('system-year-filter');
    if (systemYearFilter) {
        systemYearFilter.addEventListener('change', function() {
            const selectedYear = this.value;
            const halfYearFilter = document.getElementById('system-half-year-filter');
            const halfYearLabel = document.getElementById('system-half-year-label');
            
            // 显示或隐藏半年度筛选器
            if (halfYearFilter && halfYearLabel) {
                if (selectedYear === 'all') {
                    halfYearFilter.style.display = 'none';
                    halfYearLabel.style.display = 'none';
                    halfYearFilter.value = 'all';
                } else {
                    halfYearFilter.style.display = 'inline-block';
                    halfYearLabel.style.display = 'inline-block';
                }
            }
            
            const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
            updateSystemChart(selectedYear, selectedHalfYear);
        });
    }
    
    const systemHalfYearFilter = document.getElementById('system-half-year-filter');
    if (systemHalfYearFilter) {
        systemHalfYearFilter.addEventListener('change', function() {
            const selectedYear = document.getElementById('system-year-filter').value;
            const selectedHalfYear = this.value;
            console.log('系统图表半年度筛选器变化:', selectedYear, selectedHalfYear);
            updateSystemChart(selectedYear, selectedHalfYear);
        });
    }
    
    const typeYearFilter = document.getElementById('type-year-filter');
    if (typeYearFilter) {
        typeYearFilter.addEventListener('change', function() {
            const selectedYear = this.value;
            const halfYearFilter = document.getElementById('type-half-year-filter');
            const halfYearLabel = document.getElementById('type-half-year-label');
            
            // 显示或隐藏半年度筛选器
            if (halfYearFilter && halfYearLabel) {
                if (selectedYear === 'all') {
                    halfYearFilter.style.display = 'none';
                    halfYearLabel.style.display = 'none';
                    halfYearFilter.value = 'all';
                } else {
                    halfYearFilter.style.display = 'inline-block';
                    halfYearLabel.style.display = 'inline-block';
                }
            }
            
            const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
            updateTypeChart(selectedYear, selectedHalfYear);
        });
    }
    
    const typeHalfYearFilter = document.getElementById('type-half-year-filter');
    if (typeHalfYearFilter) {
        typeHalfYearFilter.addEventListener('change', function() {
            const selectedYear = document.getElementById('type-year-filter').value;
            const selectedHalfYear = this.value;
            updateTypeChart(selectedYear, selectedHalfYear);
        });
    }
    
    const statusYearFilter = document.getElementById('status-year-filter');
    if (statusYearFilter) {
        statusYearFilter.addEventListener('change', function() {
            const selectedYear = this.value;
            const halfYearFilter = document.getElementById('status-half-year-filter');
            const halfYearLabel = document.getElementById('status-half-year-label');
            
            // 显示或隐藏半年度筛选器
            if (halfYearFilter && halfYearLabel) {
                if (selectedYear === 'all') {
                    halfYearFilter.style.display = 'none';
                    halfYearLabel.style.display = 'none';
                    halfYearFilter.value = 'all';
                } else {
                    halfYearFilter.style.display = 'inline-block';
                    halfYearLabel.style.display = 'inline-block';
                }
            }
            
            const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
            updateStatusChart(selectedYear, selectedHalfYear);
        });
    }
    
    const statusHalfYearFilter = document.getElementById('status-half-year-filter');
    if (statusHalfYearFilter) {
        statusHalfYearFilter.addEventListener('change', function() {
            const selectedYear = document.getElementById('status-year-filter').value;
            const selectedHalfYear = this.value;
            updateStatusChart(selectedYear, selectedHalfYear);
        });
    }
    
    const monthlyYearFilter = document.getElementById('monthly-year-filter');
    if (monthlyYearFilter) {
        monthlyYearFilter.addEventListener('change', function() {
            const selectedYear = this.value;
            const halfYearFilter = document.getElementById('monthly-half-year-filter');
            const halfYearLabel = document.getElementById('monthly-half-year-label');
            
            // 显示或隐藏半年度筛选器
            if (halfYearFilter && halfYearLabel) {
                if (selectedYear === 'all') {
                    halfYearFilter.style.display = 'none';
                    halfYearLabel.style.display = 'none';
                    halfYearFilter.value = 'all';
                } else {
                    halfYearFilter.style.display = 'inline-block';
                    halfYearLabel.style.display = 'inline-block';
                }
            }
            
            const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
            updateMonthlyChart(selectedYear, selectedHalfYear);
        });
    }
    
    const monthlyHalfYearFilter = document.getElementById('monthly-half-year-filter');
    if (monthlyHalfYearFilter) {
        monthlyHalfYearFilter.addEventListener('change', function() {
            const selectedYear = document.getElementById('monthly-year-filter').value;
            const selectedHalfYear = this.value;
            updateMonthlyChart(selectedYear, selectedHalfYear);
        });
    }
    
    // 草稿勾选框事件监听器
    document.getElementById('exclude-draft').addEventListener('change', function() {
        const yearFilter = document.getElementById('status-year-filter');
        const halfYearFilter = document.getElementById('status-half-year-filter');
        const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
        updateStatusChart(yearFilter.value, selectedHalfYear);
    });
    
    document.getElementById('exclude-draft-dept').addEventListener('change', function() {
        const yearFilter = document.getElementById('dept-year-filter');
        const halfYearFilter = document.getElementById('dept-half-year-filter');
        const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
        updateDepartmentChart(yearFilter.value, selectedHalfYear);
    });
    
    // 显示原始部门切换事件监听器
    document.getElementById('show-original-dept').addEventListener('change', function() {
        const yearFilter = document.getElementById('dept-year-filter');
        const halfYearFilter = document.getElementById('dept-half-year-filter');
        const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
        updateDepartmentChart(yearFilter.value, selectedHalfYear);
    });
    
    document.getElementById('exclude-draft-system').addEventListener('change', function() {
        const yearFilter = document.getElementById('system-year-filter');
        const halfYearFilter = document.getElementById('system-half-year-filter');
        const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
        updateSystemChart(yearFilter.value, selectedHalfYear);
    });
    
    document.getElementById('exclude-draft-type').addEventListener('change', function() {
        const yearFilter = document.getElementById('type-year-filter');
        const halfYearFilter = document.getElementById('type-half-year-filter');
        const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
        updateTypeChart(yearFilter.value, selectedHalfYear);
    });
    
    document.getElementById('exclude-draft-year').addEventListener('change', function() {
        updateYearChart();
    });
    
    document.getElementById('exclude-draft-monthly').addEventListener('change', function() {
        const yearFilter = document.getElementById('monthly-year-filter');
        const halfYearFilter = document.getElementById('monthly-half-year-filter');
        const selectedHalfYear = halfYearFilter ? halfYearFilter.value : 'all';
        updateMonthlyChart(yearFilter.value, selectedHalfYear);
    });
    
    // 部门标题点击事件
    document.getElementById('dept-title').addEventListener('click', function() {
        showDepartmentModal();
    });
    
    // 模态框关闭事件
    document.getElementById('close-modal').addEventListener('click', function() {
        hideDepartmentModal();
    });
    
    // 点击模态框外部关闭
    document.getElementById('dept-modal').addEventListener('click', function(e) {
        if (e.target === this) {
            hideDepartmentModal();
        }
    });
    
    // 模态框内的筛选器事件
    const modalYearFilter = document.getElementById('modal-year-filter');
    if (modalYearFilter) {
        modalYearFilter.addEventListener('change', function() {
            const selectedYear = this.value;
            const halfYearFilter = document.getElementById('modal-half-year-filter');
            const halfYearLabel = document.getElementById('modal-half-year-label');
            
            // 显示或隐藏半年度筛选器
            if (halfYearFilter && halfYearLabel) {
                if (selectedYear === 'all') {
                    halfYearFilter.style.display = 'none';
                    halfYearLabel.style.display = 'none';
                    halfYearFilter.value = 'all';
                } else {
                    halfYearFilter.style.display = 'inline-block';
                    halfYearLabel.style.display = 'inline-block';
                }
            }
            
            updateDepartmentModalData();
        });
    }
    
    const modalHalfYearFilter = document.getElementById('modal-half-year-filter');
    if (modalHalfYearFilter) {
        modalHalfYearFilter.addEventListener('change', function() {
            updateDepartmentModalData();
        });
    }
    
    const modalExcludeDraft = document.getElementById('modal-exclude-draft');
    if (modalExcludeDraft) {
        modalExcludeDraft.addEventListener('change', function() {
            updateDepartmentModalData();
        });
    }
    
    const modalShowOriginalDept = document.getElementById('modal-show-original-dept');
    if (modalShowOriginalDept) {
        modalShowOriginalDept.addEventListener('change', function() {
            updateDepartmentModalData();
        });
    }
    
    // 原始部门切换也需要更新模态框数据
    document.getElementById('show-original-dept').addEventListener('change', function() {
        updateDepartmentModalData();
    });
    
    // ESC键关闭模态框
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            hideDepartmentModal();
        }
    });
}

// 更新筛选统计信息
function updateFilteredStats(year, halfYear = 'all') {
    const filteredCountElement = document.getElementById('filtered-count');
    
    if (year === 'all') {
        filteredCountElement.textContent = `显示全部 ${ticketData.summary.total_tickets.toLocaleString()} 张工单`;
    } else {
        const yearIndex = ticketData.year_stats.labels.indexOf(year);
        let yearCount = yearIndex >= 0 ? ticketData.year_stats.data[yearIndex] : 0;
        
        // 根据半年度筛选调整显示文本和数量
        let displayText = '';
        if (halfYear === 'all') {
            displayText = `${year}年共 ${yearCount.toLocaleString()} 张工单`;
        } else {
            // 使用月度数据计算准确的半年度统计
            const halfYearCount = calculateHalfYearCount(year, halfYear);
            if (halfYear === 'first') {
                displayText = `${year}年上半年共 ${halfYearCount.toLocaleString()} 张工单`;
            } else if (halfYear === 'second') {
                displayText = `${year}年下半年共 ${halfYearCount.toLocaleString()} 张工单`;
            }
        }
        
        filteredCountElement.textContent = displayText;
    }
}

// 根据月度数据计算半年度工单数量
function calculateHalfYearCount(year, halfYear) {
    // 获取该年的月度数据
    const monthlyData = ticketData.monthly_by_year[year];
    if (!monthlyData || !monthlyData.labels || !monthlyData.data) {
        return 0;
    }
    
    let count = 0;
    
    // 遍历月度数据，根据月份判断是上半年还是下半年
    for (let i = 0; i < monthlyData.labels.length; i++) {
        const monthLabel = monthlyData.labels[i];
        const monthValue = monthlyData.data[i];
        
        // 提取月份数字（格式：YYYY-MM）
        const month = parseInt(monthLabel.split('-')[1]);
        
        if (halfYear === 'first' && month >= 1 && month <= 6) {
            // 上半年：1-6月
            count += monthValue;
        } else if (halfYear === 'second' && month >= 7 && month <= 12) {
            // 下半年：7-12月
            count += monthValue;
        }
    }
    
    return count;
}

// 根据月度数据计算真实的半年度数据
function calculateRealHalfYearData(year, halfYear) {
    if (!ticketData.monthly_by_year || !ticketData.monthly_by_year[year]) {
        return null;
    }
    
    const monthlyData = ticketData.monthly_by_year[year];
    const labels = monthlyData.labels;
    const data = monthlyData.data;
    
    if (halfYear === 'all') {
        return { labels, data };
    }
    
    // 计算上半年（1-6月）或下半年（7-12月）的总数
    let total = 0;
    for (let i = 0; i < labels.length; i++) {
        const month = parseInt(labels[i].split('-')[1]); // 提取月份
        if (halfYear === 'first' && month >= 1 && month <= 6) {
            total += data[i];
        } else if (halfYear === 'second' && month >= 7 && month <= 12) {
            total += data[i];
        }
    }
    
    return total;
}

// 计算半年度部门数据
function calculateHalfYearDepartmentData(year, halfYear, excludeDraft, isOriginal) {
    // 获取年度总数据
    const sourceData = isOriginal ? 
        (excludeDraft ? ticketData.original_dept_by_year_all_no_draft : ticketData.original_dept_by_year_all) :
        (excludeDraft ? ticketData.dept_by_year_no_draft : ticketData.dept_by_year);
    
    const yearData = sourceData[year] || { labels: [], data: [] };
    
    if (halfYear === 'all') {
        return yearData;
    }
    
    // 获取该年度的真实半年度比例
    const yearTotal = calculateRealHalfYearData(year, 'all');
    const halfYearTotal = calculateRealHalfYearData(year, halfYear);
    
    if (!yearTotal || !halfYearTotal || yearTotal.data.reduce((a, b) => a + b, 0) === 0) {
        // 如果没有月度数据，使用默认比例
        const ratio = halfYear === 'first' ? 0.6 : 0.4;
        return {
            labels: yearData.labels,
            data: yearData.data.map(count => Math.round(count * ratio))
        };
    }
    
    // 计算真实比例
    const totalYearTickets = yearTotal.data.reduce((a, b) => a + b, 0);
    const ratio = halfYearTotal / totalYearTickets;
    
    return {
        labels: yearData.labels,
        data: yearData.data.map(count => Math.round(count * ratio))
    };
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
function updateDepartmentChart(year, halfYear = 'all') {
    if (!charts.department) return;
    
    const excludeDraft = document.getElementById('exclude-draft-dept').checked;
    const showOriginal = document.getElementById('show-original-dept').checked;
    let data;
    
    if (showOriginal) {
        // 显示原始部门数据
        if (year === 'all') {
            const sourceData = excludeDraft ? ticketData.original_dept_all_no_draft : ticketData.original_dept_all;
            // 取前10个部门显示
            data = {
                labels: sourceData.labels.slice(0, 10),
                data: sourceData.data.slice(0, 10)
            };
        } else {
            const sourceData = excludeDraft ? ticketData.original_dept_by_year_all_no_draft : ticketData.original_dept_by_year_all;
            let yearData = sourceData[year] || { labels: [], data: [] };
            
            // 如果选择了半年度筛选，需要根据月度数据重新计算
            if (halfYear !== 'all' && ticketData.monthly_by_year && ticketData.monthly_by_year[year]) {
                yearData = calculateHalfYearDepartmentData(year, halfYear, excludeDraft, true);
            }
            
            // 取前10个部门显示
            data = {
                labels: yearData.labels.slice(0, 10),
                data: yearData.data.slice(0, 10)
            };
        }
    } else {
        // 显示映射后的一级部门数据
        if (year === 'all') {
            data = excludeDraft ? ticketData.dept_top10_no_draft : ticketData.dept_top10;
        } else {
            const sourceData = excludeDraft ? ticketData.dept_by_year_no_draft : ticketData.dept_by_year;
            let yearData = sourceData[year] || { labels: [], data: [] };
            
            // 如果选择了半年度筛选，需要根据月度数据重新计算
            if (halfYear !== 'all' && ticketData.monthly_by_year && ticketData.monthly_by_year[year]) {
                yearData = calculateHalfYearDepartmentData(year, halfYear, excludeDraft, false);
            }
            
            data = yearData;
        }
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
function updateSystemChart(year, halfYear = 'all') {
    if (!charts.system) return;
    
    const excludeDraft = document.getElementById('exclude-draft-system').checked;
    let data;
    
    if (year === 'all') {
        data = excludeDraft ? ticketData.system_stats_no_draft : ticketData.system_stats;
    } else {
        const sourceData = excludeDraft ? ticketData.system_by_year_no_draft : ticketData.system_by_year;
        let yearData = sourceData[year] || { 'OA系统': 0, '营销平台': 0, 'U8C': 0 };
        
        // 如果选择了半年度筛选，使用真实的月度数据计算比例
        if (halfYear !== 'all') {
            // 获取该年度的真实半年度比例
            const yearTotal = calculateRealHalfYearData(year, 'all');
            const halfYearTotal = calculateRealHalfYearData(year, halfYear);
            
            let ratio;
            if (yearTotal && halfYearTotal && yearTotal.data.reduce((a, b) => a + b, 0) > 0) {
                // 使用真实比例
                const totalYearTickets = yearTotal.data.reduce((a, b) => a + b, 0);
                ratio = halfYearTotal / totalYearTickets;
            } else {
                // 如果没有月度数据，使用默认比例
                ratio = halfYear === 'first' ? 0.6 : 0.4;
            }
            
            const adjustedData = {};
            Object.keys(yearData).forEach(key => {
                adjustedData[key] = Math.round(yearData[key] * ratio);
            });
            data = adjustedData;
        } else {
            data = yearData;
        }
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

// 更新年度趋势图表
function updateYearChart() {
    if (!charts.year) return;
    
    const excludeDraft = document.getElementById('exclude-draft-year').checked;
    let data;
    
    if (excludeDraft) {
        data = ticketData.year_stats_no_draft || ticketData.year_stats;
    } else {
        data = ticketData.year_stats;
    }
    
    charts.year.data.labels = data.labels;
    charts.year.data.datasets[0].data = data.data;
    charts.year.update();
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
function updateTypeChart(year, halfYear = 'all') {
    if (!charts.type) return;
    
    const excludeDraft = document.getElementById('exclude-draft-type').checked;
    let data;
    
    if (year === 'all') {
        data = excludeDraft ? ticketData.type_stats_no_draft : ticketData.type_stats;
    } else {
        const sourceData = excludeDraft ? ticketData.type_by_year_no_draft : ticketData.type_by_year;
        let yearData = sourceData[year] || { labels: [], data: [] };
        
        // 如果选择了半年度筛选，使用真实的月度数据计算比例
        if (halfYear !== 'all') {
            // 获取该年度的真实半年度比例
            const yearTotal = calculateRealHalfYearData(year, 'all');
            const halfYearTotal = calculateRealHalfYearData(year, halfYear);
            
            let ratio;
            if (yearTotal && halfYearTotal && yearTotal.data.reduce((a, b) => a + b, 0) > 0) {
                // 使用真实比例
                const totalYearTickets = yearTotal.data.reduce((a, b) => a + b, 0);
                ratio = halfYearTotal / totalYearTickets;
            } else {
                // 如果没有月度数据，使用默认比例
                ratio = halfYear === 'first' ? 0.6 : 0.4;
            }
            
            data = {
                labels: yearData.labels,
                data: yearData.data.map(count => Math.round(count * ratio))
            };
        } else {
            data = yearData;
        }
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
            onClick: function(event, elements) {
                if (elements.length > 0) {
                    const elementIndex = elements[0].index;
                    const label = statusData.labels[elementIndex];
                    
                    // 如果点击的是"未结束"状态，跳转到详情页面
                    if (label === '未结束') {
                        window.open('details.html', '_blank');
                    }
                }
            },
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
                            const label = context.label;
                            const value = context.parsed.y;
                            let tooltip = `${value} 张工单`;
                            if (label === '未结束') {
                                tooltip += ' (点击查看详情)';
                            }
                            return tooltip;
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
function updateStatusChart(year, halfYear = 'all') {
    if (!charts.status) return;
    
    let statusData;
    let auditData;
    
    if (year === 'all') {
        statusData = ticketData.status_stats;
        auditData = ticketData.audit_stats;
    } else {
        statusData = ticketData.status_by_year[year] || { labels: [], data: [] };
        auditData = ticketData.audit_by_year[year] || { labels: [], data: [] };
    }
    
    // 如果选择了半年度筛选，使用真实的月度数据计算比例
    if (halfYear !== 'all' && year !== 'all') {
        // 获取该年度的真实半年度比例
        const yearTotal = calculateRealHalfYearData(year, 'all');
        const halfYearTotal = calculateRealHalfYearData(year, halfYear);
        
        let ratio;
        if (yearTotal && halfYearTotal && yearTotal.data.reduce((a, b) => a + b, 0) > 0) {
            // 使用真实比例
            const totalYearTickets = yearTotal.data.reduce((a, b) => a + b, 0);
            ratio = halfYearTotal / totalYearTickets;
        } else {
            // 如果没有月度数据，使用默认比例
            ratio = halfYear === 'first' ? 0.6 : 0.4;
        }
        
        statusData = {
            labels: statusData.labels,
            data: statusData.data.map(value => Math.round(value * ratio))
        };
        auditData = {
            labels: auditData.labels,
            data: auditData.data.map(value => Math.round(value * ratio))
        };
    }
    
    // 检查是否需要排除草稿
    const excludeDraft = document.getElementById('exclude-draft').checked;
    
    let filteredLabels = [];
    let filteredData = [];
    
    if (excludeDraft && auditData.labels.includes('草稿')) {
        // 获取草稿数量
        const draftIndex = auditData.labels.indexOf('草稿');
        const draftCount = auditData.data[draftIndex];
        
        // 从状态数据中减去草稿数量（草稿通常在"未结束"状态中）
        for (let i = 0; i < statusData.labels.length; i++) {
            const label = statusData.labels[i];
            let count = statusData.data[i];
            
            // 如果是"未结束"状态，减去草稿数量
            if (label === '未结束') {
                count = Math.max(0, count - draftCount);
            }
            
            // 只添加数量大于0的状态
            if (count > 0) {
                filteredLabels.push(label);
                filteredData.push(count);
            }
        }
    } else {
        // 不过滤，使用原始数据
        filteredLabels = statusData.labels;
        filteredData = statusData.data;
    }
    
    charts.status.data.labels = filteredLabels;
    charts.status.data.datasets[0].data = filteredData;
    charts.status.data.datasets[0].backgroundColor = chartColors.slice(0, filteredLabels.length);
    charts.status.data.datasets[0].borderColor = chartColors.slice(0, filteredLabels.length);
    
    // 重新设置点击事件
    charts.status.options.onClick = function(event, elements) {
        if (elements.length > 0) {
            const elementIndex = elements[0].index;
            const label = filteredLabels[elementIndex];
            
            // 如果点击的是"未结束"状态，跳转到详情页面
            if (label === '未结束') {
                window.open('details.html', '_blank');
            }
        }
    };
    
    // 重新设置tooltip
    charts.status.options.plugins.tooltip.callbacks.label = function(context) {
        const label = context.label;
        const value = context.parsed.y;
        let tooltip = `${value} 张工单`;
        if (label === '未结束') {
            tooltip += ' (点击查看详情)';
        }
        return tooltip;
    };
    
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
function updateMonthlyChart(year, halfYear = 'all') {
    if (!charts.monthly) return;
    
    const excludeDraft = document.getElementById('exclude-draft-monthly').checked;
    let data;
    
    if (year === 'all') {
        data = excludeDraft ? ticketData.monthly_stats_no_draft : ticketData.monthly_stats;
    } else {
        const sourceData = excludeDraft ? ticketData.monthly_by_year_no_draft : ticketData.monthly_by_year;
        data = sourceData[year] || { labels: [], data: [] };
    }
    
    // 如果选择了半年度筛选，只显示对应的月份数据
    if (halfYear !== 'all' && year !== 'all') {
        const monthIndices = halfYear === 'first' ? [0, 1, 2, 3, 4, 5] : [6, 7, 8, 9, 10, 11];
        const filteredLabels = [];
        const filteredData = [];
        
        monthIndices.forEach(index => {
            if (index < data.labels.length) {
                filteredLabels.push(data.labels[index]);
                filteredData.push(data.data[index] || 0);
            }
        });
        
        data = {
            labels: filteredLabels,
            data: filteredData
        };
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

// 显示部门详情模态框
function showDepartmentModal() {
    const modal = document.getElementById('dept-modal');
    
    // 初始化模态框的年度筛选器
    initializeModalYearFilter();
    
    // 显示模态框
    modal.style.display = 'block';
    
    // 加载部门数据
    updateDepartmentModalData();
    
    // 防止背景滚动
    document.body.style.overflow = 'hidden';
}

// 隐藏部门详情模态框
function hideDepartmentModal() {
    const modal = document.getElementById('dept-modal');
    modal.style.display = 'none';
    
    // 恢复背景滚动
    document.body.style.overflow = 'auto';
}

// 初始化模态框的年度筛选器
function initializeModalYearFilter() {
    const modalYearFilter = document.getElementById('modal-year-filter');
    
    // 清空现有选项
    modalYearFilter.innerHTML = '<option value="all">全部年份</option>';
    
    // 添加年份选项
    if (ticketData && ticketData.dept_by_year) {
        const years = Object.keys(ticketData.dept_by_year).sort();
        years.forEach(year => {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = year + '年';
            modalYearFilter.appendChild(option);
        });
    }
}

// 更新模态框中的部门数据
function updateDepartmentModalData() {
    if (!ticketData) return;
    
    const year = document.getElementById('modal-year-filter').value;
    const excludeDraft = document.getElementById('modal-exclude-draft').checked;
    const tableBody = document.getElementById('dept-table-body');
    
    // 获取完整的部门数据
    let allDeptData = getAllDepartmentData(year, excludeDraft);
    
    // 清空表格
    tableBody.innerHTML = '';
    
    // 填充数据
    allDeptData.forEach((dept, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>${dept.name}</td>
            <td>${dept.count}</td>
        `;
        tableBody.appendChild(row);
    });
    
    // 如果没有数据，显示提示
    if (allDeptData.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="3" style="text-align: center; color: #666; padding: 40px;">暂无部门数据</td>';
        tableBody.appendChild(row);
    }
}

// 获取所有部门数据（完整列表，按工单数量排序）
function getAllDepartmentData(year, excludeDraft) {
    if (!ticketData) return [];
    
    // 检查模态框内的原始部门复选框状态
    const modalShowOriginal = document.getElementById('modal-show-original-dept').checked;
    let sourceData;
    
    if (modalShowOriginal) {
        // 显示原始部门数据
        if (year === 'all') {
            sourceData = excludeDraft ? ticketData.original_dept_all_no_draft : ticketData.original_dept_all;
        } else {
            const yearData = excludeDraft ? ticketData.original_dept_by_year_all_no_draft : ticketData.original_dept_by_year_all;
            sourceData = yearData[year] || { labels: [], data: [] };
        }
    } else {
        // 显示映射后的一级部门数据
        if (year === 'all') {
            sourceData = excludeDraft ? ticketData.dept_all_no_draft : ticketData.dept_all;
        } else {
            const yearData = excludeDraft ? ticketData.dept_by_year_all_no_draft : ticketData.dept_by_year_all;
            sourceData = yearData[year] || { labels: [], data: [] };
        }
    }
    
    // 转换为数组格式
    const sortedDepts = sourceData.labels.map((name, index) => ({
        name: name,
        count: sourceData.data[index]
    }));
    
    return sortedDepts;
}