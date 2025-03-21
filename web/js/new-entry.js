// New Entry page JavaScript

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
    
    // Set current date
    const currentDateInput = document.getElementById('currentDate');
    if(currentDateInput) {
        const today = new Date();
        const formattedDate = today.toISOString().split('T')[0];
        currentDateInput.value = formattedDate;
    }
    
    // Check if we're editing an existing site
    const urlParams = new URLSearchParams(window.location.search);
    const siteId = urlParams.get('site_id');
    
    if(siteId) {
        // Fetch site data
        fetchSiteData(siteId);
    }
    
    // Agreement Date and Lease Period calculation
    const agreementDateInput = document.getElementById('agreementDate');
    const leasePeriodSelect = document.getElementById('leasePeriod');
    const agreementValidUptoInput = document.getElementById('agreementValidUpto');
    
    function calculateAgreementValidUpto() {
        if(!agreementDateInput || !leasePeriodSelect || !agreementValidUptoInput) return;
        
        const agreementDate = agreementDateInput.value;
        const leasePeriod = parseInt(leasePeriodSelect.value || "0", 10);
        
        if(!agreementDate || isNaN(leasePeriod)) return;
        
        const date = new Date(agreementDate);
        date.setFullYear(date.getFullYear() + leasePeriod);
        
        agreementValidUptoInput.value = date.toISOString().split('T')[0];
        calculateValidityDate();
    }
    
    if(agreementDateInput && leasePeriodSelect) {
        agreementDateInput.addEventListener('change', calculateAgreementValidUpto);
        leasePeriodSelect.addEventListener('change', calculateAgreementValidUpto);
    }
    
    // Rent Position Date and Current Date calculation
    const rentPositionDateInput = document.getElementById('rentPositionDate');
    const currentDate1Input = document.getElementById('currentDate1');
    
    function calculateCurrentDate1() {
        if(!rentPositionDateInput || !currentDateInput || !currentDate1Input) return;
        
        const rentPositionDate = rentPositionDateInput.value;
        const currentDate = currentDateInput.value;
        
        if(!rentPositionDate || !currentDate) return;
        
        const rpDate = new Date(rentPositionDate);
        const cdDate = new Date(currentDate);
        
        const yearDiff = cdDate.getFullYear() - rpDate.getFullYear();
        const monthDiff = cdDate.getMonth() - rpDate.getMonth();
        const dayDiff = cdDate.getDate() - rpDate.getDate();
        
        let years = yearDiff;
        let months = monthDiff;
        let days = dayDiff;
        
        if(dayDiff < 0) {
            months--;
            const lastMonth = new Date(cdDate.getFullYear(), cdDate.getMonth(), 0);
            days = lastMonth.getDate() + dayDiff;
        }
        
        if(months < 0) {
            years--;
            months += 12;
        }
        
        currentDate1Input.value = `${years} Years, ${months} Months, ${days} Days`;
    }
    
    if(rentPositionDateInput) {
        rentPositionDateInput.addEventListener('change', calculateCurrentDate1);
    }
    
    // Agreement Valid Upto and Current Date calculation
    const validityDateInput = document.getElementById('validityDate');
    
    function calculateValidityDate() {
        if(!agreementValidUptoInput || !currentDateInput || !validityDateInput) return;
        
        const agreementValidUpto = agreementValidUptoInput.value;
        const currentDate = currentDateInput.value;
        
        if(!agreementValidUpto || !currentDate) return;
        
        const avuDate = new Date(agreementValidUpto);
        const cdDate = new Date(currentDate);
        
        const yearDiff = avuDate.getFullYear() - cdDate.getFullYear();
        const monthDiff = avuDate.getMonth() - cdDate.getMonth();
        const dayDiff = avuDate.getDate() - cdDate.getDate();
        
        let years = yearDiff;
        let months = monthDiff;
        let days = dayDiff;
        
        if(dayDiff < 0) {
            months--;
            const lastMonth = new Date(avuDate.getFullYear(), avuDate.getMonth(), 0);
            days = lastMonth.getDate() + dayDiff;
        }
        
        if(months < 0) {
            years--;
            months += 12;
        }
        
        validityDateInput.value = `${years} Years, ${months} Months, ${days} Days`;
    }
    
    // Upload Excel button and modal
    const uploadExcelBtn = document.getElementById('uploadExcel');
    const uploadModal = document.getElementById('uploadModal');
    const closeModalBtn = uploadModal ? uploadModal.querySelector('.close') : null;
    const cancelUploadBtn = document.getElementById('cancelUpload');
    
    if(uploadExcelBtn && uploadModal) {
        uploadExcelBtn.addEventListener('click', function() {
            uploadModal.style.display = 'flex';
        });
    }
    
    if(closeModalBtn) {
        closeModalBtn.addEventListener('click', function() {
            uploadModal.style.display = 'none';
        });
    }
    
    if(cancelUploadBtn) {
        cancelUploadBtn.addEventListener('click', function() {
            uploadModal.style.display = 'none';
        });
    }
    
    // Close modal when clicking outside
    window.addEventListener('click', function(e) {
        if(e.target === uploadModal) {
            uploadModal.style.display = 'none';
        }
    });
    
    // Upload Form
    const uploadForm = document.getElementById('uploadForm');
    
    if(uploadForm) {
        uploadForm.addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const fileInput = document.getElementById('excelFile');
            
            if(!fileInput || !fileInput.files.length) {
                alert('Please select a file to upload');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            
            try {
                const response = await fetch('http://localhost:5000/api/upload', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${token}`
                    },
                    body: formData
                });
                
                const data = await response.json();
                
                if(response.ok) {
                    alert('File uploaded successfully');
                    uploadModal.style.display = 'none';
                } else {
                    alert(data.message || 'Upload failed');
                }
            } catch(error) {
                console.error('Upload error:', error);
                alert('An error occurred during upload');
            }
        });
    }
    
    // New Entry Form
    const newEntryForm = document.getElementById('newEntryForm');
    
    if(newEntryForm) {
        newEntryForm.addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const formData = new FormData(newEntryForm);
            const data = {};
            
            for(const [key, value] of formData.entries()) {
                data[key] = value;
            }
            
            try {
                const url = siteId ? 
                    `http://localhost:5000/api/sites/${siteId}` : 
                    'http://localhost:5000/api/sites';
                
                const response = await fetch(url, {
                    method: siteId ? 'PUT' : 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${token}`
                    },
                    body: JSON.stringify(data)
                });
                
                const result = await response.json();
                
                if(response.ok) {
                    alert(siteId ? 'Site updated successfully' : 'Site created successfully');
                    window.location.href = 'dashboard.html';
                } else {
                    alert(result.message || 'Operation failed');
                }
            } catch(error) {
                console.error('Form submission error:', error);
                alert('An error occurred while saving the data');
            }
        });
    }
    
    // Clear Form Button
    const clearFormBtn = document.getElementById('clearFormBtn');
    
    if(clearFormBtn && newEntryForm) {
        clearFormBtn.addEventListener('click', function() {
            const confirmClear = confirm('Are you sure you want to clear the form?');
            
            if(confirmClear) {
                newEntryForm.reset();
                
                // Reset the current date after form reset
                if(currentDateInput) {
                    const today = new Date();
                    const formattedDate = today.toISOString().split('T')[0];
                    currentDateInput.value = formattedDate;
                }
            }
        });
    }
    
    // Cancel Button
    const cancelBtn = document.getElementById('cancelBtn');
    
    if(cancelBtn) {
        cancelBtn.addEventListener('click', function() {
            const confirmCancel = confirm('Are you sure you want to cancel? Any unsaved data will be lost.');
            
            if(confirmCancel) {
                window.location.href = 'dashboard.html';
            }
        });
    }
    
    // Report Link
    const reportLink = document.getElementById('reportLink');
    
    if(reportLink) {
        reportLink.addEventListener('click', function(e) {
            e.preventDefault();
            window.location.href = 'reports.html';
        });
    }
    
    // Fetch site data for editing
    async function fetchSiteData(siteId) {
        try {
            const response = await fetch(`http://localhost:5000/api/sites?site_id=${siteId}`, {
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            });
            
            const data = await response.json();
            
            if(response.ok) {
                // Populate form fields
                const site = data.site;
                for(const [key, value] of Object.entries(site)) {
                    const input = document.getElementById(key);
                    if(input) {
                        input.value = value;
                    }
                }
                
                // Update calculated fields
                calculateAgreementValidUpto();
                calculateCurrentDate1();
                calculateValidityDate();
            } else {
                alert(data.message || 'Failed to fetch site data');
            }
        } catch(error) {
            console.error('Fetch error:', error);
            alert('An error occurred while fetching site data');
        }
    }
}); 