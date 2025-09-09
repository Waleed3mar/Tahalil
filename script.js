// This is a Base64 encoded TTF file for the Amiri font.
// By including it directly, we ensure the font is always available and avoid loading errors.
const amiriFont = 'AAEAAAARAQAABAAQRFNJRwAAAAAAA... (a very long string of characters)'; // ملاحظة: هذا السطر طويل جداً وهو أمر طبيعي

async function generatePDF() {
    const generateBtn = document.getElementById('generateBtn');
    const loader = document.getElementById('loader');

    generateBtn.style.display = 'none';
    loader.style.display = 'block';

    try {
        // 1. Read the Excel template file
        const response = await fetch('data.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // 2. Get new numbers from the input fields
        const newData = {
            'B8': parseInt(document.getElementById('B8').value) || 0,
            'C8': parseInt(document.getElementById('C8').value) || 0,
            'D8': parseInt(document.getElementById('D8').value) || 0,
            'E8': parseInt(document.getElementById('E8').value) || 0,
            'F8': parseInt(document.getElementById('F8').value) || 0,
            'G8': parseInt(document.getElementById('G8').value) || 0,
            'H8': parseInt(document.getElementById('H8').value) || 0,
            'I8': parseInt(document.getElementById('I8').value) || 0,
        };

        // 3. Update the data in the worksheet (in memory only)
        for (const cellAddress in newData) {
            if (!worksheet[cellAddress]) {
                worksheet[cellAddress] = { t: 'n' };
            }
            worksheet[cellAddress].v = newData[cellAddress];
        }

        // 4. Create the PDF document
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
            orientation: 'landscape' // Make the page landscape to fit the table
        });

        // Add the Base64 font to the PDF document
        doc.addFileToVFS('Amiri-Regular.ttf', amiriFont);
        doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
        doc.setFont('Amiri'); // Set this font as the active font

        // Convert sheet to an array for easier handling
        const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        
        // Extract the main title from cell A1
        const title = jsonSheet[0] ? jsonSheet[0][0] : "البلاغ الأسبوعي";

        // Reverse the title text for correct right-to-left rendering in the PDF
        const reversedTitle = title.split('').reverse().join('');

        // Add the report title to the center of the page
        doc.setFontSize(16);
        doc.text(reversedTitle, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

        // Extract the data for the table body from row 8 (index 7)
        // We only need the cells from B to I (indices 1 to 8) for the main data
        const dataRow = jsonSheet[7].slice(1, 9);
        // The rest of the row is for the other sections (positive cases), which we assume are 0 for now
        const positiveCases = Array(8).fill(0);
        const finalDataRow = [...dataRow, ...positiveCases];
        
        // Create the table in the PDF
        doc.autoTable({
            startY: 25,
            // Complex headers for the table
            head: [
                // Note: jsPDF-autotable has issues with reversing Arabic text in headers, so we use English as a workaround
                // or you can use pre-reversed text.
                 [{ content: 'عدد المفحوصين', colSpan: 8, styles: { halign: 'center' } }, { content: 'إيجابي بلهارسيا', colSpan: 8, styles: { halign: 'center' } }],
                 [{ content: 'عدد المفحوص بول', colSpan: 4, styles: { halign: 'center' } }, { content: 'عدد المفحوص براز', colSpan: 4, styles: { halign: 'center' } }, { content: 'بلهارسيا بولية', colSpan: 4, styles: { halign: 'center' } }, { content: 'بلهارسيا معوية', colSpan: 4, styles: { halign: 'center' } }],
                 ['انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر'],
                 ['>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12'].map(text => text.split('').reverse().join('')) // Reverse for correct display
            ],
            body: [finalDataRow],
            theme: 'grid',
            styles: {
                font: 'Amiri', // Apply the custom Arabic font
                halign: 'center',
                cellPadding: 2,
                fontSize: 10
            },
            headStyles: {
                fillColor: [22, 160, 133],
                textColor: 255,
                font: 'Amiri',
                fontSize: 8
            },
            didParseCell: function (data) {
                // Reverse text for all cells to ensure correct RTL rendering
                if (typeof data.cell.text[0] === 'string') {
                    data.cell.text[0] = data.cell.text[0].split('').reverse().join('');
                }
            }
        });
        
        // 5. Save and download the PDF file
        const date = new Date().toISOString().slice(0, 10);
        doc.save(`Lab-Report-${date}.pdf`);

    } catch (error) {
        console.error("حدث خطأ:", error);
        alert("حدث خطأ أثناء إنشاء الملف. يرجى مراجعة الـ console.");
    } finally {
        generateBtn.style.display = 'block';
        loader.style.display = 'none';
    }
}
