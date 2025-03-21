// Reports page JavaScript

document.addEventListener('DOMContentLoaded', function() {
    // Check authentication
    const token = localStorage.getItem('token');
    if (!token) {
        window.location.href = 'index.html';
        return;
    }

    // Toggle sidebar
    const toggleBtn = document.querySelector('.toggle-btn');
    const sidebar = document.querySelector('.sidebar');
    
    if(toggleBtn && sidebar) {
        toggleBtn.addEventListener('click', function() {
            sidebar.classList.toggle('active');
        });
    }
    
    // New Entry Link
    const newEntryLink = document.getElementById('newEntryLink');
    
    if(newEntryLink) {
        newEntryLink.addEventListener('click', function(e) {
            e.preventDefault();
            window.location.href = 'new-entry.html';
        });
    }
    
    // Report Type Radio Buttons
    const reportTypeRadios = document.querySelectorAll('input[name="reportType"]');
    const additionalFilters = document.getElementById('additionalFilters');
    const leasePeriodFilter = document.getElementById('leasePeriodFilter');
    
    if(reportTypeRadios.length && additionalFilters && leasePeriodFilter) {
        reportTypeRadios.forEach(radio => {
            radio.addEventListener('change', function() {
                const reportType = this.value;
                
                // Show/hide additional filters based on report type
                if(reportType === 'ALL SITES DATA REPORTS') {
                    additionalFilters.style.display = 'flex';
                } else {
                    additionalFilters.style.display = 'none';
                }
                
                // Show/hide lease period filter based on report type
                if(reportType === 'Lease Period Report') {
                    leasePeriodFilter.style.display = 'block';
                } else {
                    leasePeriodFilter.style.display = 'none';
                }
            });
        });
    }
    
    // Generate Report Button
    const generateBtn = document.getElementById('generateBtn');
    
    if(generateBtn) {
        generateBtn.addEventListener('click', async function() {
            const fromDate = document.getElementById('fromDate').value;
            const toDate = document.getElementById('toDate').value;
            
            if(!fromDate || !toDate) {
                alert('Please select both from and to dates');
                return;
            }
            
            const selectedReportType = document.querySelector('input[name="reportType"]:checked').value;
            
            // Check additional inputs based on report type
            if(selectedReportType === 'Lease Period Report') {
                const leasePeriod = document.getElementById('leasePeriod').value;
                if(!leasePeriod) {
                    alert('Please select a lease period');
                    return;
                }
            }
            
            try {
                const url = new URL('http://localhost:5000/api/reports');
                url.searchParams.append('type', selectedReportType);
                url.searchParams.append('from_date', fromDate);
                url.searchParams.append('to_date', toDate);
                
                if(selectedReportType === 'Lease Period Report') {
                    url.searchParams.append('lease_period', document.getElementById('leasePeriod').value);
                }
                
                const response = await fetch(url, {
                    headers: {
                        'Authorization': `Bearer ${token}`
                    }
                });
                
                const data = await response.json();
                
                if(response.ok) {
                    displayReportData(data.data, selectedReportType);
                } else {
                    alert(data.message || 'Failed to generate report');
                }
            } catch(error) {
                console.error('Report generation error:', error);
                alert('An error occurred while generating the report');
            }
        });
    }
    
    // Back Button
    const backBtn = document.getElementById('backBtn');
    
    if(backBtn) {
        backBtn.addEventListener('click', function() {
            window.location.href = 'dashboard.html';
        });
    }
    
    // Search in Report
    const reportSearch = document.getElementById('reportSearch');
    
    if(reportSearch) {
        reportSearch.addEventListener('input', function() {
            const searchTerm = this.value.toLowerCase();
            const tableRows = document.querySelectorAll('#reportTable tbody tr');
            
            tableRows.forEach(row => {
                const text = row.textContent.toLowerCase();
                if(text.includes(searchTerm)) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        });
    }
    
    // Export to Excel Button
    const exportBtn = document.querySelector('.btn-export');
    
    if(exportBtn) {
        exportBtn.addEventListener('click', function() {
            const table = document.getElementById('reportTable');
            if(!table) return;
            
            // Create CSV content
            let csv = [];
            const rows = table.querySelectorAll('tr');
            
            rows.forEach(row => {
                const rowData = [];
                row.querySelectorAll('th, td').forEach(cell => {
                    rowData.push(cell.textContent);
                });
                csv.push(rowData.join(','));
            });
            
            // Create and download file
            const blob = new Blob([csv.join('\n')], { type: 'text/csv' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'report.csv';
            a.click();
            window.URL.revokeObjectURL(url);
        });
    }
    
    // Display report data in table
    function displayReportData(data, reportType) {
        const tableHead = document.querySelector('#reportTable thead tr');
        const tableBody = document.querySelector('#reportTable tbody');
        
        if(!tableHead || !tableBody) return;
        
        // Clear existing content
        tableHead.innerHTML = '';
        tableBody.innerHTML = '';
        
        // Add headers
        const headers = Object.keys(data[0] || {});
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.replace(/_/g, ' ').toUpperCase();
            tableHead.appendChild(th);
        });
        
        // Add data rows
        data.forEach(row => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = row[header];
                tr.appendChild(td);
            });
            tableBody.appendChild(tr);
        });
    }
}); 