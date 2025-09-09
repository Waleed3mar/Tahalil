function prepareAndPrint() {
    // 1. Get values from input fields
    const values = {
        B8: document.getElementById('B8').value || 0,
        C8: document.getElementById('C8').value || 0,
        D8: document.getElementById('D8').value || 0,
        E8: document.getElementById('E8').value || 0,
        F8: document.getElementById('F8').value || 0,
        G8: document.getElementById('G8').value || 0,
        H8: document.getElementById('H8').value || 0,
        I8: document.getElementById('I8').value || 0,
    };

    // 2. Update the hidden report table with the new values
    for (const key in values) {
        document.getElementById(`report-${key}`).innerText = values[key];
    }
    
    // 3. Open the browser's print dialog
    window.print();
}
