/**
 * Admin Dashboard - Interactive JavaScript
 * Handles chart rendering, tooltips, sidebar interactions
 */

// ============================================
// Sidebar Interactions
// ============================================

function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (sidebar) {
        sidebar.classList.toggle('collapsed');
    }
}

function toggleSubmenu(element) {
    element.classList.toggle('expanded');
    const submenu = element.nextElementSibling;
    if (submenu && submenu.classList.contains('nav-submenu')) {
        submenu.classList.toggle('show');
    }
}

// ============================================
// Bar Chart Rendering with Hover Tooltips
// ============================================

const barChartData = {
    months: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    revenue: [45000, 52000, 48000, 61000, 55000, 67000, 72000, 69000, 78000, 82000, 79000, 91000],
    grossMargin: [32000, 38000, 35000, 44000, 40000, 49000, 53000, 50000, 57000, 60000, 58000, 67000]
};

function renderBarChart() {
    const container = document.getElementById('barChart');
    if (!container) return;

    const maxValue = Math.max(...barChartData.revenue, ...barChartData.grossMargin);
    const chartHeight = 200;
    
    let html = '<div class="bar-chart">';
    
    barChartData.months.forEach((month, index) => {
        const revenueHeight = (barChartData.revenue[index] / maxValue) * chartHeight;
        const marginHeight = (barChartData.grossMargin[index] / maxValue) * chartHeight;
        
        html += `
            <div class="bar-group" data-month="${month}" data-revenue="${barChartData.revenue[index]}" data-margin="${barChartData.grossMargin[index]}">
                <div class="bar bar-indigo" style="height: ${revenueHeight}px;" data-value="$${(barChartData.revenue[index]/1000).toFixed(0)}K"></div>
                <div class="bar bar-amber" style="height: ${marginHeight}px;" data-value="$${(barChartData.grossMargin[index]/1000).toFixed(0)}K"></div>
            </div>
        `;
    });
    
    html += '</div>';
    
    // X-axis labels
    html += '<div class="chart-x-axis">';
    barChartData.months.forEach(month => {
        html += `<span class="x-label">${month}</span>`;
    });
    html += '</div>';
    
    container.innerHTML = html;
    
    // Add hover event listeners for tooltips
    setupBarChartTooltips(container);
}

function setupBarChartTooltips(container) {
    const tooltip = document.getElementById('chartTooltip');
    const barGroups = container.querySelectorAll('.bar-group');
    
    barGroups.forEach(group => {
        group.addEventListener('mouseenter', (e) => {
            const month = group.dataset.month;
            const revenue = group.dataset.revenue;
            const margin = group.dataset.margin;
            
            tooltip.querySelector('.tooltip-month').textContent = month;
            tooltip.querySelector('.tooltip-values').innerHTML = `
                <div class="tooltip-value-item">
                    <span class="tooltip-dot" style="background: #5A6ACF;"></span>
                    <span class="tooltip-label">Revenue</span>
                    <span class="tooltip-value">$${(parseInt(revenue)/1000).toFixed(0)}K</span>
                </div>
                <div class="tooltip-value-item">
                    <span class="tooltip-dot" style="background: #F59E0B;"></span>
                    <span class="tooltip-label">Gross Margin</span>
                    <span class="tooltip-value">$${(parseInt(margin)/1000).toFixed(0)}K</span>
                </div>
            `;
            
            tooltip.classList.add('visible');
        });
        
        group.addEventListener('mousemove', (e) => {
            const rect = tooltip.getBoundingClientRect();
            const tooltipWidth = rect.width;
            const tooltipHeight = rect.height;
            
            let left = e.clientX + 15;
            let top = e.clientY - tooltipHeight - 10;
            
            // Prevent tooltip from going off-screen
            if (left + tooltipWidth > window.innerWidth) {
                left = e.clientX - tooltipWidth - 15;
            }
            if (top < 0) {
                top = e.clientY + 15;
            }
            
            tooltip.style.left = `${left}px`;
            tooltip.style.top = `${top}px`;
        });
        
        group.addEventListener('mouseleave', () => {
            tooltip.classList.remove('visible');
        });
    });
}

// ============================================
// Initialize on DOM Ready
// ============================================

document.addEventListener('DOMContentLoaded', function() {
    // Render charts
    renderBarChart();
    
    // Setup icon rail active state
    const appIcons = document.querySelectorAll('.app-icon');
    appIcons.forEach(icon => {
        icon.addEventListener('click', function() {
            appIcons.forEach(i => i.classList.remove('active'));
            this.classList.add('active');
        });
    });
    
    // Setup nav item active state
    const navItems = document.querySelectorAll('.nav-item:not(.has-submenu)');
    navItems.forEach(item => {
        item.addEventListener('click', function(e) {
            if (!this.classList.contains('has-submenu')) {
                navItems.forEach(i => i.classList.remove('active'));
                this.classList.add('active');
            }
        });
    });
    
    // Search bar keyboard shortcut
    document.addEventListener('keydown', function(e) {
        if ((e.metaKey || e.ctrlKey) && e.key === 'f') {
            e.preventDefault();
            const searchInput = document.querySelector('.search-input');
            if (searchInput) {
                searchInput.focus();
            }
        }
    });
    
    // Card view more buttons
    const viewMoreButtons = document.querySelectorAll('.card-view-more');
    viewMoreButtons.forEach(btn => {
        btn.addEventListener('click', function() {
            // In a real app, this would navigate to a detail page
            console.log('View more clicked');
        });
    });
    
    // Hero CTA button
    const heroCta = document.querySelector('.hero-cta');
    if (heroCta) {
        heroCta.addEventListener('click', function() {
            console.log('Explore features clicked');
        });
    }
    
    // Table action buttons (event delegation)
    const tableCard = document.querySelector('.table-card');
    if (tableCard) {
        tableCard.addEventListener('click', function(e) {
            const actionBtn = e.target.closest('.action-btn');
            if (actionBtn) {
                const action = actionBtn.title;
                console.log(`${action} clicked`);
            }
            
            const pageBtn = e.target.closest('.page-btn');
            if (pageBtn && !pageBtn.classList.contains('active')) {
                document.querySelectorAll('.page-btn').forEach(btn => btn.classList.remove('active'));
                pageBtn.classList.add('active');
            }
        });
    }
    
    // Activity list add member button
    const addMemberBtn = document.querySelector('.card-action-btn');
    if (addMemberBtn) {
        addMemberBtn.addEventListener('click', function() {
            console.log('Add member clicked');
        });
    }
    
    // Table action buttons
    const filterBtn = document.querySelector('.filter-btn');
    const exportBtn = document.querySelector('.export-btn');
    const addBtn = document.querySelector('.add-btn');
    
    if (filterBtn) filterBtn.addEventListener('click', () => console.log('Filter clicked'));
    if (exportBtn) exportBtn.addEventListener('click', () => console.log('Export clicked'));
    if (addBtn) addBtn.addEventListener('click', () => console.log('Add member clicked'));
});

// ============================================
// Utility Functions
// ============================================

// Format currency
function formatCurrency(value) {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(value);
}

// Format percentage
function formatPercentage(value) {
    return `${value >= 0 ? '+' : ''}${value.toFixed(1)}%`;
}

// Debounce function for resize events
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Handle window resize
window.addEventListener('resize', debounce(function() {
    // Re-render charts if needed on significant size changes
    const container = document.getElementById('barChart');
    if (container) {
        renderBarChart();
    }
}, 250));
