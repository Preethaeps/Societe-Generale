document.addEventListener("DOMContentLoaded", () => {
    const dropArea = document.getElementById("drop-area");
    const fileInput = document.getElementById("file-input");
    const fileBtn = document.getElementById("file-btn");
    const output = document.getElementById("output");

    dropArea.addEventListener("dragover", (event) => {
        event.preventDefault();
        dropArea.classList.add("dragover");
    });

    dropArea.addEventListener("dragleave", () => {
        dropArea.classList.remove("dragover");
    });

    dropArea.addEventListener("drop", (event) => {
        event.preventDefault();
        dropArea.classList.remove("dragover");
        const files = event.dataTransfer.files;
        handleFiles(files);
    });

    fileBtn.addEventListener("click", () => {
        fileInput.click();
    });

    fileInput.addEventListener("change", (event) => {
        const files = event.target.files;
        handleFiles(files);
    });

    function handleFiles(files) {
        const file = files[0];
        if (file && (file.name.endsWith(".xlsm") || file.name.endsWith(".xlsx"))) {
            const formData = new FormData();
            formData.append("file", file);

            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    output.innerHTML = `<p>${data.error}</p>`;
                } else {
                    output.innerHTML = "<h3>Analysis Result:</h3>";
                    data.forEach(macro => {
                        output.innerHTML += `<p><strong>Filename:</strong> ${macro.filename}</p>
                                             <p><strong>Stream Path:</strong> ${macro.stream_path}</p>
                                             <p><strong>VBA Filename:</strong> ${macro.vba_filename}</p>
                                             <pre><strong>VBA Code:</strong>\n${macro.vba_code}</pre>`;
                    });
                }
            })
            .catch(error => {
                output.innerHTML = `<p>An error occurred: ${error.message}</p>`;
            });
        } else {
            output.innerHTML = "<p>Please upload a valid Excel file with VBA Macros (.xlsm or .xlsx).</p>";
        }
    }
});
