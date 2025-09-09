// This is a Base64 encoded TTF file for the Amiri font.
// It is very long, which is normal. Make sure it is copied completely.
const amiriFont =
  'AAEAAAARAQAABAAQRFNJRwAAAAAAA...'; // This string is intentionally truncated here for display, but ensure the full string is in your actual file.

async function generatePDF() {
  const generateBtn = document.getElementById('generateBtn');
  const loader = document.getElementById('loader');

  generateBtn.style.display = 'none';
  loader.style.display = 'block';

  try {
    // 1. Read the Excel template file
    const response = await fetch('data.xlsx');
    if (!response.ok) {
      throw new Error('Failed to load the Excel template file (data.xlsx). Make sure the file exists in your project.');
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 2. Get new numbers from the input fields
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

    // 3. Update the data in the worksheet (in memory only)
    for (const cellAddress in newData) {
      if (!worksheet[cellAddress]) worksheet[cellAddress] = { t: 'n' };
      // THIS LINE IS NOW CORRECTED
      worksheet[cellAddress].v = newData[cellAddress];
    }

    // 4. Create the PDF document
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape' });

    // Add the embedded font to the PDF document
    doc.addFileToVFS('Amiri-Regular.ttf', amiriFont);
    doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
    doc.setFont('Amiri');

    // Convert sheet to an array for easier handling
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    // Extract the main title and render it correctly
    const title = String(jsonSheet[0]?.[0] || 'البلاغ الأسبوعي');
    doc.setFontSize(16);
    // jsPDF requires manual text reversal for right-to-left languages
    const reversedTitle = title.split('').reverse().join('');
    doc.text(reversedTitle, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });
    
    // Extract the updated data row
    const dataRow = jsonSheet[7].slice(1, 17);

    // Create the table in the PDF
    doc.autoTable({
      startY: 25,
      head: [
        ['إيجابي بلهارسيا', 'عدد المفحوصين'].map(t => t.split('').reverse().join('')),
        ['معوية', 'بولية', 'براز', 'بول'].map(t => t.split('').reverse().join('')),
        ['ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى'],
        ['>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12']
      ],
      body: [dataRow],
      theme: 'grid',
      styles: {
        font: 'Amiri',
        halign: 'center',
        cellPadding: 2,
        fontSize: 10,
      },
      headStyles: {
        fillColor: [41, 128, 185],
        textColor: 255,
        font: 'Amiri',
        fontSize: 10,
        valign: 'middle',
      },
      // This function ensures all text in the table is rendered correctly from right-to-left
      didParseCell: function (data) {
          // Process only body cells to avoid reversing numbers and symbols
          if (data.section === 'body') {
            if (typeof data.cell.text[0] === 'string') {
              // You might not need to reverse numbers, so check if the content is a number
              if (isNaN(data.cell.text[0])) {
                  data.cell.text[0] = data.cell.text[0].split('').reverse().join('');
              }
            }
          }
      }
    });

    // 5. Save and download the PDF file
    const date = new Date().toISOString().slice(0, 10);
    doc.save(`Lab-Report-${date}.pdf`);
  } catch (error) {
    console.error('An error occurred:', error);
    alert('An error occurred while creating the file. Please check the console for details.');
  } finally {
    generateBtn.style.display = 'block';
    loader.style.display = 'none';
  }
}
