/* ============================================
   Shared UI JavaScript
   Sidebar interactions and common utilities
   ============================================ */

// Toggle sidebar collapse
function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (sidebar) {
        sidebar.classList.toggle('collapsed');
        
        // Save state to localStorage
        const isCollapsed = sidebar.classList.contains('collapsed');
        localStorage.setItem('sidebarCollapsed', isCollapsed);
    }
}

// Toggle submenu expansion
function toggleSubmenu(element) {
    element.classList.toggle('expanded');
    const submenu = element.nextElementSibling;
    if (submenu && submenu.classList.contains('nav-submenu')) {
        submenu.classList.toggle('show');
    }
}

// Restore sidebar state on page load
document.addEventListener('DOMContentLoaded', function() {
    const sidebar = document.getElementById('sidebar');
    if (sidebar) {
        const isCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
        if (isCollapsed) {
            sidebar.classList.add('collapsed');
        }
    }
    
    // Set active nav item based on current page
    setActiveNavItem();
});

// Set active navigation item based on current page
function setActiveNavItem() {
    const currentPage = window.location.pathname.split('/').pop() || 'index.html';
    const navItems = document.querySelectorAll('.nav-item');
    
    navItems.forEach(item => {
        const href = item.getAttribute('href');
        if (href === currentPage || (currentPage === '' && href === 'index.html')) {
            item.classList.add('active');
        } else {
            item.classList.remove('active');
        }
    });
}

// Chart tooltip interactions
function initChartTooltips() {
    const bars = document.querySelectorAll('.bar');
    const tooltip = document.querySelector('.chart-tooltip');
    
    if (!tooltip) return;
    
    bars.forEach(bar => {
        bar.addEventListener('mouseenter', function(e) {
            const value = this.getAttribute('data-value');
            const label = this.getAttribute('data-label');
            
            tooltip.innerHTML = `<strong>${label}</strong><br>${value}`;
            tooltip.classList.add('show');
        });
        
        bar.addEventListener('mousemove', function(e) {
            const rect = this.getBoundingClientRect();
            tooltip.style.left = `${rect.left + rect.width / 2 - tooltip.offsetWidth / 2}px`;
            tooltip.style.top = `${rect.top - tooltip.offsetHeight - 10}px`;
        });
        
        bar.addEventListener('mouseleave', function() {
            tooltip.classList.remove('show');
        });
    });
}

// Initialize tooltips when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    initChartTooltips();
});
