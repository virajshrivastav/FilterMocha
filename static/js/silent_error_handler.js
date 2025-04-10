/**
 * Silent Error Handler
 * 
 * This script adds error handling to the frontend that suppresses error popups
 * and instead logs errors silently to the console.
 */

document.addEventListener('DOMContentLoaded', function() {
    // Override the default fetch error handling
    const originalFetch = window.fetch;
    window.fetch = function() {
        return originalFetch.apply(this, arguments)
            .then(response => {
                if (!response.ok) {
                    // Log the error but don't throw it
                    console.error(`HTTP error: ${response.status} ${response.statusText}`);
                    
                    // For API endpoints, return a valid response structure instead of throwing
                    if (arguments[0].includes('/api/')) {
                        // Create a fake successful response with empty data
                        return response.json().catch(() => {
                            // If we can't parse JSON, return a minimal structure
                            return {
                                silent_error: `HTTP error: ${response.status} ${response.statusText}`,
                                success: false
                            };
                        });
                    }
                }
                return response;
            })
            .catch(error => {
                // Log network errors but don't throw them for API calls
                console.error('Network error:', error);
                
                if (arguments[0].includes('/api/')) {
                    // Return a minimal valid response structure
                    return {
                        silent_error: `Network error: ${error.message}`,
                        success: false
                    };
                }
                
                // Re-throw for non-API calls
                throw error;
            });
    };
    
    // Override the default error handling in the app.js
    if (window.showError) {
        const originalShowError = window.showError;
        window.showError = function(message) {
            // Log the error to console instead of showing it
            console.error('Application error (suppressed):', message);
            
            // Don't call the original function to avoid showing the error popup
            // originalShowError(message);
        };
    }
    
    // Add a global error handler for AJAX requests
    $(document).ajaxError(function(event, jqXHR, settings, thrownError) {
        // Log the error but don't show it
        console.error('AJAX error (suppressed):', thrownError);
        console.error('URL:', settings.url);
        console.error('Status:', jqXHR.status);
        console.error('Response:', jqXHR.responseText);
        
        // Prevent default error handling
        event.preventDefault();
        event.stopPropagation();
        
        // Return false to prevent further handling
        return false;
    });
    
    // Check for silent errors in API responses and handle them gracefully
    const originalAjax = $.ajax;
    $.ajax = function(options) {
        const originalSuccess = options.success;
        
        if (originalSuccess) {
            options.success = function(data) {
                // Check if the response contains a silent error
                if (data && data.silent_error) {
                    // Log the error but don't show it
                    console.error('Silent API error:', data.silent_error);
                    
                    // If there's output data, still process it
                    if (data.output_files && data.output_files.length > 0) {
                        originalSuccess(data);
                    } else if (data.columns !== undefined) {
                        // For analyze endpoint, still process the response
                        originalSuccess(data);
                    } else {
                        // Create a minimal valid response
                        const minimalData = {
                            success: false,
                            output_files: [],
                            errors: []
                        };
                        originalSuccess(minimalData);
                    }
                } else {
                    // Normal processing
                    originalSuccess(data);
                }
            };
        }
        
        return originalAjax.apply(this, [options]);
    };
    
    console.log('Silent error handler initialized');
});
