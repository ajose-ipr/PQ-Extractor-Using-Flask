// static/js/main.js

document.addEventListener('DOMContentLoaded', function() {
    // File upload progress indicator
    const uploadForm = document.getElementById('upload-form');
    if (uploadForm) {
        uploadForm.addEventListener('submit', function() {
            const submitBtn = this.querySelector('button[type="submit"]');
            submitBtn.disabled = true;
            submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Processing...';
        });
    }

    // Dynamic file selection feedback
    const fileInput = document.getElementById('file-input');
    if (fileInput) {
        fileInput.addEventListener('change', function() {
            const fileList = Array.from(this.files)
                .map(file => file.name)
                .join(', ');
            document.getElementById('file-selection-feedback').textContent = 
                fileList || 'No files selected';
        });
    }

    // Table row highlighting for violations
    document.querySelectorAll('.violation-row').forEach(row => {
        row.addEventListener('click', function() {
            this.classList.toggle('table-warning');
            const details = this.nextElementSibling;
            details.classList.toggle('d-none');
        });
    });

    // Tab functionality for voltage/current sections
    const tabLinks = document.querySelectorAll('.nav-tabs .nav-link');
    tabLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const target = this.getAttribute('data-bs-target');
            document.querySelectorAll('.tab-pane').forEach(pane => {
                pane.classList.remove('active', 'show');
            });
            document.querySelector(target).classList.add('active', 'show');
            
            tabLinks.forEach(l => l.classList.remove('active'));
            this.classList.add('active');
        });
    });
});