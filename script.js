// This is a Base64 encoded TTF file for the Amiri font.
// By including it directly, we ensure the font is always available and avoid loading errors.
// ملاحظة: هذا السطر طويل جداً وهو أمر طبيعي وصحيح. لا تقم بتعديله.
const amiriFont = 'AAEAAAARAQAABAAQRFNJRwAAAAAAA...'; // This string will be extremely long.

async function generatePDF() {
    const generateBtn = document.getElementById('generateBtn');
    const loader = document.getElementById('loader');

    generateBtn.style.display = 'none';
    loader.style.display = 'block';

    try {
        // 1. قراءة ملف الإكسيل كقالب
        const response = await fetch('data.xlsx');
        if (!response.ok) throw new Error('Network response was not ok');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // 2. قراءة الأرقام الجديدة من حقول الإدخال
        const newData = {
            'B8': parseInt(document.getElementById('B8').value, 10) || 0,
            'C8': parseInt(document.getElementById('C8').value, 10) || 0,
            'D8': parseInt(document.getElementById('D8').value, 10) || 0,
            'E8': parseInt(document.getElementById('E8').value, 10) || 0,
            'F8': parseInt(document.getElementById('F8').value, 10) || 0,
            'G8': parseInt(document.getElementById('G8').value, 10) || 0,
            'H8': parseInt(document.getElementById('H8').value, 10) || 0,
            'I8': parseInt(document.getElementById('I8').value, 10) || 0,
        };

        // 3. تحديث البيانات في الشيت (في الذاكرة فقط)
        for (const cellAddress in newData) {
            if (!worksheet[cellAddress]) worksheet[cellAddress] = { t: 'n' };
            worksheet[cellAddress].v = newData[cellAddress];
        }

        // 4. إنشاء ملف PDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'landscape' });

        // **هذا هو الجزء الذي تم تصحيحه**
        // إضافة الخط المدمج إلى المستند الافتراضي وتفعيله
        doc.addFileToVFS('Amiri-Regular.ttf', amiriFont);
        doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
        doc.setFont('Amiri');

        // تحويل الشيت إلى مصفوفة لسهولة التعامل
        const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        
        // استخراج عنوان التقرير
        const title = jsonSheet[0] ? String(jsonSheet[0][0]) : "البلاغ الأسبوعي";
        doc.setFontSize(16);
        doc.text(title, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

        // استخراج صف البيانات المحدث
        const dataRow = jsonSheet[7].slice(1, 17);

        // إنشاء الجدول في الـ PDF
        doc.autoTable({
            startY: 25,
            head: [
                // تم تبسيط الرؤوس لضمان التوافقية
                [{ content: 'عدد المفحوصين', colSpan: 8, styles: { halign: 'center', font: 'Amiri' } }, { content: 'إيجابي بلهارسيا', colSpan: 8, styles: { halign: 'center', font: 'Amiri' } }],
                [{ content: 'بول', colSpan: 4, styles: { halign: 'center', font: 'Amiri' } }, { content: 'براز', colSpan: 4, styles: { halign: 'center', font: 'Amiri' } }, { content: 'بولية', colSpan: 4, styles: { halign: 'center', font: 'Amiri' } }, { content: 'معوية', colSpan: 4, styles: { halign: 'center', font: 'Amiri' } }],
                ['انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر'],
                ['>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12']
            ],
            body: [dataRow],
            theme: 'grid',
            styles: {
                font: 'Amiri', // تطبيق الخط العربي على كل الخلايا
                halign: 'center',
                cellPadding: 2,
                fontSize: 10,
            },
            headStyles: {
                fillColor: [41, 128, 185],
                textColor: 255,
                font: 'Amiri',
                fontSize: 10
            }
        });
        
        // 5. حفظ وتنزيل ملف الـ PDF
        const date = new Date().toISOString().slice(0, 10);
        doc.save(`Lab-Report-${date}.pdf`);

    } catch (error) {
        console.error("حدث خطأ:", error);
        alert("حدث خطأ أثناء إنشاء الملف. يرجى المحاولة مرة أخرى.");
    } finally {
        generateBtn.style.display = 'block';
        loader.style.display = 'none';
    }
}
