// Dashboard page JavaScript

document.addEventListener('DOMContentLoaded', function() {
    // Check authentication
    const token = localStorage.getItem('token');
    if (!token) {
        window.location.href = 'index.html';
        return;
    }

    // Function to handle unauthorized responses
    async function handleUnauthorized() {
        alert('Your session has expired. Please log in again.');
        localStorage.removeItem('token');
        localStorage.removeItem('user');
        window.location.href = 'index.html';
    }

    // Function to make authenticated API calls
    async function makeAuthenticatedRequest(url, options = {}) {
        const token = localStorage.getItem('token');
        if (!token) {
            handleUnauthorized();
            return null;
        }

        const response = await fetch(url, {
            ...options,
            headers: {
                ...options.headers,
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });

        if (response.status === 401) {
            await handleUnauthorized();
            return null;
        }

        return response;
    }

    // Toggle sidebar
    const toggleBtn = document.querySelector('.toggle-btn');
    const sidebar = document.querySelector('.sidebar');
    
    if(toggleBtn && sidebar) {
        toggleBtn.addEventListener('click', function() {
            sidebar.classList.toggle('active');
        });
    }
    
    // Search form
    const searchForm = document.getElementById('searchForm');
    
    if(searchForm) {
        searchForm.addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const siteSearch = document.getElementById('siteSearch').value;
            
            if(!siteSearch) {
                alert('Please enter a site ID');
                return;
            }
            
            try {
                const response = await makeAuthenticatedRequest(`http://localhost:5000/api/sites?site_id=${siteSearch}`);
                if (!response) return;
                
                const data = await response.json();
                
                if(response.ok) {
                    updateSiteData(data.site);
                } else {
                    alert(data.message || 'Site not found');
                }
            } catch(error) {
                console.error('Search error:', error);
                alert('An error occurred while searching');
            }
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
    
    // Report Link
    const reportLink = document.getElementById('reportLink');
    
    if(reportLink) {
        reportLink.addEventListener('click', function(e) {
            e.preventDefault();
            window.location.href = 'reports.html';
        });
    }
    
    // Edit Button
    const editBtn = document.getElementById('editBtn');
    let currentSiteId = null; // Store current site ID
    
    if(editBtn) {
        // Create cancel button but keep it hidden initially
        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = 'Cancel';
        cancelBtn.className = 'btn btn-cancel';
        cancelBtn.style.display = 'none';
        cancelBtn.id = 'cancelBtn';
        editBtn.parentNode.insertBefore(cancelBtn, editBtn.nextSibling);
        
        editBtn.addEventListener('click', async function() {
            if (!editBtn.classList.contains('editing')) {
                // Show authentication modal first
                const authModal = document.getElementById('authModal');
                authModal.style.display = 'flex';

                // Remove any existing event listeners to prevent duplicates
                const authForm = document.getElementById('authForm');
                const newAuthForm = authForm.cloneNode(true);
                authForm.parentNode.replaceChild(newAuthForm, authForm);

                // Handle authentication form submission
                newAuthForm.addEventListener('submit', function(e) {
                    e.preventDefault();
                    const username = document.getElementById('authUsername').value;
                    const password = document.getElementById('authPassword').value;

                    if (username === 'admin' && password === 'admin@123') {
                        authModal.style.display = 'none';
                        newAuthForm.reset();
                        startEditing(); // Call the existing edit functionality
                    } else {
                        const errorMsg = document.querySelector('#authModal .error-message') || 
                            (() => {
                                const div = document.createElement('div');
                                div.className = 'error-message';
                                div.style.color = 'red';
                                div.style.marginTop = '10px';
                                newAuthForm.appendChild(div);
                                return div;
                            })();
                        errorMsg.textContent = 'Invalid credentials';
                        errorMsg.style.display = 'block';
                    }
                });

                // Handle cancel button
                const cancelAuth = document.getElementById('cancelAuth');
                cancelAuth.addEventListener('click', function() {
                    authModal.style.display = 'none';
                    newAuthForm.reset();
                    const errorMsg = document.querySelector('#authModal .error-message');
                    if (errorMsg) errorMsg.style.display = 'none';
                });

                // Handle modal close button if it exists
                const closeBtn = authModal.querySelector('.close');
                if (closeBtn) {
                    closeBtn.addEventListener('click', function() {
                        authModal.style.display = 'none';
                        newAuthForm.reset();
                        const errorMsg = document.querySelector('#authModal .error-message');
                        if (errorMsg) errorMsg.style.display = 'none';
                    });
                }
            } else {
                try {
                    // Saving changes
                    const updatedData = {};
                    const valueElements = document.querySelectorAll('[id$="_value"]');
                    
                    valueElements.forEach(element => {
                        const fieldName = element.id.replace('_value', '');
                        const input = element.querySelector('input');
                        const value = input ? input.value : element.textContent;
                        
                        if (value !== '') { // Only include non-empty values
                            updatedData[fieldName] = value.trim();
                        }
                    });

                    console.log('Data being sent to server:', updatedData);
                    
                    if (Object.keys(updatedData).length === 0) {
                        throw new Error('No data to update');
                    }

                    await updateSite(currentSiteId, updatedData);
                    
                    // Reset the edit mode
                    editBtn.textContent = 'Edit';
                    editBtn.classList.remove('editing');
                    cancelBtn.style.display = 'none';
                    
                } catch (error) {
                    console.error('Error updating site:', error);
                    alert('Error updating site: ' + error.message);
                }
            }
        });

        // Cancel button handler
        cancelBtn.addEventListener('click', async function() {
            // Reset the form to its original state
            editBtn.textContent = 'Edit';
            editBtn.classList.remove('editing');
            cancelBtn.style.display = 'none';
            
            // Refresh the display with original data
            if (currentSiteId) {
                const response = await makeAuthenticatedRequest(`http://localhost:5000/api/sites?site_id=${currentSiteId}`);
                if (!response) return;
                
                const data = await response.json();
                if (response.ok) {
                    updateSiteData(data.site);
                }
            }
        });
    }
    
    // Function to update site data on the server
    async function updateSite(siteId, data) {
        try {
            const response = await makeAuthenticatedRequest(`http://localhost:5000/api/sites/${siteId}`, {
                method: 'PUT',
                body: JSON.stringify(data)
            });
            if (!response) return;
            
            const result = await response.json();
            
            if (response.ok) {
                alert('Site updated successfully');
                // Refresh the display with updated data
                const searchResponse = await makeAuthenticatedRequest(`http://localhost:5000/api/sites?site_id=${siteId}`);
                if (!searchResponse) return;
                
                const searchData = await searchResponse.json();
                if (searchResponse.ok) {
                    updateSiteData(searchData.site);
                }
            } else {
                alert(result.message || 'Error updating site');
            }
        } catch (error) {
            console.error('Update error:', error);
            alert('An error occurred while updating the site');
        }
    }
    
    // Logout
    const logoutLink = document.querySelector('a[href="index.html"]');
    if(logoutLink) {
        logoutLink.addEventListener('click', function(e) {
            e.preventDefault();
            localStorage.removeItem('token');
            localStorage.removeItem('user');
            window.location.href = 'index.html';
        });
    }
    
    // Update site data in the UI
    function updateSiteData(siteData) {
        console.log("Raw site data received:", siteData);
        
        if (!siteData) {
            console.error("No site data provided");
            return;
        }
        
        // Helper function to safely update element text content
        function updateElement(id, value) {
            const element = document.getElementById(id);
            if (element) {
                console.log(`Updating element #${id} with value:`, value);
                element.textContent = value !== null && value !== undefined ? value : '';
            } else {
                console.warn(`Element with ID ${id} not found`);
            }
        }

        // Check for specific fields
        console.log("Site ID:", siteData.site);
        console.log("CURRENT DATE 1:", siteData.current_date1);
        console.log("VALIDITY DATE:", siteData.validity_date);
        console.log("HIKE %:", siteData.hike_percentage);
        
        // Update site ID with both possible IDs to ensure it's displayed
        updateElement('site_value', siteData.site);
        updateElement('site_id_value', siteData.site);
        
        // Update basic info elements
        updateElement('store_name_value', siteData.store_name);
        updateElement('region_value', siteData.region);
        updateElement('div_value', siteData.div);
        updateElement('manager_value', siteData.manager);
        updateElement('asst_manager_value', siteData.asst_manager);
        updateElement('executive_value', siteData.executive);
        updateElement('doo_value', siteData.doo);
        updateElement('sqft_value', siteData.sqft);
        updateElement('agreement_date_value', siteData.agreement_date);
        updateElement('rent_position_date_value', siteData.rent_position_date);
        updateElement('rent_effective_date_value', siteData.rent_effective_date);
        updateElement('agreement_valid_upto_value', siteData.agreement_valid_upto);
        updateElement('current_date_value', siteData.current_date);
        updateElement('lease_period_value', siteData.lease_period);
        updateElement('rent_free_period_days_value', siteData.rent_free_period_days);
        updateElement('rent_effective_amount_value', siteData.rent_effective_amount);
        updateElement('present_rent_value', siteData.present_rent);
        
        // Special handling for hike percentage to ensure it's displayed with the % symbol
        if (siteData.hike_percentage !== null && siteData.hike_percentage !== undefined) {
            let hikeValue = siteData.hike_percentage;
            try {
                // Parse as float if it's a string
                if (typeof hikeValue === 'string') {
                    hikeValue = parseFloat(hikeValue);
                }
                
                // Format the value with the % symbol if it's a number
                if (!isNaN(hikeValue)) {
                    hikeValue = hikeValue + '%';
                }
            } catch (e) {
                console.error("Error formatting hike percentage:", e);
            }
            updateElement('hike_percentage_value', hikeValue);
        } else {
            updateElement('hike_percentage_value', '');
        }
        
        // Update remaining elements
        updateElement('hike_year_value', siteData.hike_year);
        updateElement('rent_deposit_value', siteData.rent_deposit);
        updateElement('owner_name1_value', siteData.owner_name1);
        updateElement('owner_name2_value', siteData.owner_name2);
        updateElement('owner_name3_value', siteData.owner_name3);
        updateElement('owner_name4_value', siteData.owner_name4);
        updateElement('owner_name5_value', siteData.owner_name5);
        updateElement('owner_name6_value', siteData.owner_name6);
        updateElement('owner_mobile_value', siteData.owner_mobile);
        
        // Directly use the calculated values for CURRENT DATE 1 and VALIDITY DATE
        // These values are already calculated in the backend and formatted
        updateElement('current_date1_value', siteData.current_date1);
        updateElement('validity_date_value', siteData.validity_date);
        
        updateElement('gst_number_value', siteData.gst_number);
        updateElement('pan_number_value', siteData.pan_number);
        updateElement('tds_percentage_value', siteData.tds_percentage);
        updateElement('mature_value', siteData.mature);
        updateElement('status_value', siteData.status);
        updateElement('remarks_value', siteData.remarks);

        // Make sure the site details section is visible
        const siteDetailsSection = document.getElementById('site_details_section');
        if (siteDetailsSection) {
            siteDetailsSection.style.display = 'block';
        } else {
            console.warn("Site details section not found");
        }
    }

    // Helper function to format currency values
    function formatCurrency(value) {
        if (!value) return '';
        return new Intl.NumberFormat('en-IN', {
            style: 'currency',
            currency: 'INR',
            maximumFractionDigits: 0
        }).format(value);
    }

    // Helper function to format dates
    function formatDate(dateString) {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString; // Return original string if invalid date
        return date.toLocaleDateString('en-IN');
    }

    // Add this helper function at the top level
    function formatDateForDisplay(dateStr) {
        if (!dateStr) return '';
        try {
            // Handle both yyyy-mm-dd and dd-mm-yyyy formats
            let parts;
            if (dateStr.includes('-')) {
                parts = dateStr.split('-');
                if (parts[0].length === 4) {
                    // yyyy-mm-dd format
                    return `${parts[2]}-${parts[1]}-${parts[0]}`;
                } else {
                    // dd-mm-yyyy format
                    return dateStr;
                }
            }
            return dateStr;
        } catch (e) {
            console.error('Date formatting error:', e);
            return dateStr;
        }
    }

    // Add this function to handle the edit mode start
    function startEditing() {
        // Starting edit mode
        let siteElement = document.getElementById('site_value');
        // If the site_value element isn't found, try site_id_value
        if (!siteElement) {
            siteElement = document.getElementById('site_id_value');
        }
        
        currentSiteId = siteElement ? siteElement.textContent : null;
        console.log("Current site ID for editing:", currentSiteId);
        
        if(!currentSiteId) {
            alert('No site selected');
            return;
        }

        // Change button text to "Save" and show cancel button
        editBtn.textContent = 'Save';
        editBtn.classList.add('editing');
        cancelBtn.style.display = 'inline-block';
        
        // Make all value fields editable
        const valueElements = document.querySelectorAll('[id$="_value"]');
        valueElements.forEach(element => {
            const fieldName = element.id.replace('_value', '');
            // Check if this is a date field
            if (fieldName.includes('date') || fieldName === 'doo') {
                const currentValue = element.textContent;
                // Create date input
                const dateInput = document.createElement('input');
                dateInput.type = 'text';
                dateInput.value = formatDateForDisplay(currentValue);
                dateInput.className = 'form-control';
                dateInput.placeholder = 'dd-mm-yyyy';
                element.textContent = '';
                element.appendChild(dateInput);
            } else {
                const input = document.createElement('input');
                input.type = 'text';
                input.value = element.textContent;
                input.className = 'form-control';
                element.textContent = '';
                element.appendChild(input);
            }
        });
    }
});