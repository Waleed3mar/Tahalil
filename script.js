// This is the Base64 encoded Amiri font. It is very long, which is normal.
// Please ensure this entire line is copied correctly without any modification.
const amiriFont =
  'AAEAAAARAQAABAAQRFNJRwAAAAAAA...'; // IMPORTANT: This string is intentionally truncated for display. The full version is required for the code to work.

async function generatePDF() {
  const generateBtn = document.getElementById('generateBtn');
  const loader = document.getElementById('loader');

  generateBtn.style.display = 'none';
  loader.style.display = 'block';

  try {
    // STEP 1: Fetch and Read the Excel File
    console.log('Step 1: Attempting to fetch data.xlsx...');
    const response = await fetch('data.xlsx');
    if (!response.ok) {
      // This error will be shown if the file is not found (404 error)
      throw new Error(`Failed to load data.xlsx. Status: ${response.status}. Please make sure the file name is exactly 'data.xlsx' and it is in the same directory as index.html.`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    console.log('Step 1: Excel file loaded successfully.');

    // STEP 2: Get New Values from Input Fields
    const newData = {
      B8: parseInt(document.getElementById('B8').value, 10) || 0,
      C8: parseInt(document.getElementById('C8').value, 10) || 0,
      D8: parseInt(document.getElementById('D8').value, 10) || 0,
      E8: parseInt(document.getElementById('E8').value, 10) || 0,
      F8: parseInt(document.getElementById('F8').value, 10) || 0,
      G8: parseInt(document.getElementById('G8').value, 10) || 0,
      H8: parseInt(document.getElementById('H8').value, 10) || 0,
      I8: parseInt(document.getElementById('I8').value, 10) || 0,
    };

    // STEP 3: Update Worksheet Data in Memory
    for (const cellAddress in newData) {
      if (!worksheet[cellAddress]) worksheet[cellAddress] = { t: 'n' };
      worksheet[cellAddress].v = newData[cellAddress];
    }

    // STEP 4: Generate the PDF Document
    console.log('Step 4: Creating PDF document...');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape' });

    // Add the embedded font to the PDF
    doc.addFileToVFS('Amiri-Regular.ttf', amiriFont);
    doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
    doc.setFont('Amiri');

    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const title = String(jsonSheet[0]?.[0] || 'البلاغ الأسبوعي');
    const dataRow = jsonSheet[7].slice(1, 17);

    const reversedTitle = title.split('').reverse().join('');
    doc.setFontSize(16);
    doc.text(reversedTitle, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

    doc.autoTable({
      startY: 25,
      head: [
          [{ content: 'ايجابى بلهارسيا'.split('').reverse().join(''), colSpan: 8, styles: { halign: 'center' } }, { content: 'عدد المفحوصين'.split('').reverse().join(''), colSpan: 8, styles: { halign: 'center' } }],
          [{ content: 'معوية'.split('').reverse().join(''), colSpan: 4, styles: { halign: 'center' } }, { content: 'بولية'.split('').reverse().join(''), colSpan: 4, styles: { halign: 'center' } }, { content: 'براز'.split('').reverse().join(''), colSpan: 4, styles: { halign: 'center' } }, { content: 'بول'.split('').reverse().join(''), colSpan: 4, styles: { halign: 'center' } }],
          ['ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى'].map(t => t.split('').reverse().join('')),
          ['>12', '<=12', '>12', '<=12', '>12', '<=12', '>12', '<=12', '>12', '<=12', '>12', '<=12', '>12', '<=12', '>12', '<=12']
      ],
      body: [dataRow],
      theme: 'grid',
      styles: { font: 'Amiri', halign: 'center', cellPadding: 2, fontSize: 10, },
      headStyles: { fillColor: [41, 128, 185], textColor: 255, font: 'Amiri', fontSize: 9, valign: 'middle' }
    });
    
    // STEP 5: Save the PDF
    console.log('Step 5: Saving PDF...');
    const date = new Date().toISOString().slice(0, 10);
    doc.save(`Lab-Report-${date}.pdf`);

  } catch (error) {
    // This new error handling is more robust.
    console.error("A critical error occurred. Details below:");
    console.error(error); // This will print the full error object, whatever it is.
    alert("An error occurred: " + (error.message || "Unknown error. Check the console (F12) for more details."));
  } finally {
    generateBtn.style.display = 'block';
    loader.style.display = 'none';
  }
}
