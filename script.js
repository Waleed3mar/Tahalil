async function generatePDF() {
    const generateBtn = document.getElementById('generateBtn');
    const loader = document.getElementById('loader');

    // إظهار رسالة التحميل وإخفاء الزر
    generateBtn.style.display = 'none';
    loader.style.display = 'block';

    try {
        // 1. قراءة ملف الإكسيل كقالب
        const response = await fetch('data.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // 2. قراءة الأرقام الجديدة من حقول الإدخال
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

        // 3. تحديث البيانات في الشيت (في الذاكرة فقط)
        for (const cellAddress in newData) {
            // إضافة الخلية إذا لم تكن موجودة
            if (!worksheet[cellAddress]) {
                worksheet[cellAddress] = { t: 'n', v: 0 };
            }
            worksheet[cellAddress].v = newData[cellAddress];
        }

        // 4. إنشاء ملف PDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
            orientation: 'landscape' // لجعل الصفحة بالعرض لتناسب الجدول
        });

        // تحويل الشيت إلى مصفوفة لسهولة التعامل معها في pdf
        const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        // استخراج البيانات المهمة لإنشاء الجدول
        const title = jsonSheet[0][0]; // عنوان التقرير من الخلية A1
        const headers1 = jsonSheet[3].slice(1, 17); // رؤوس الجدول من الصف الرابع
        const headers2 = jsonSheet[4].slice(1, 17);
        const headers3 = jsonSheet[5].slice(1, 17);
        const headers4 = jsonSheet[6].slice(1, 17);
        const dataRow = jsonSheet[7].slice(1, 17); // صف البيانات من الصف الثامن
        
        // لإضافة الخطوط العربية (اختياري لكن يحسن الشكل)
        doc.addFont('https://cdnjs.cloudflare.com/ajax/libs/js-pdf/1.3.2/Amiri-Regular-normal.js', 'Amiri', 'normal');
        doc.setFont('Amiri');

        // إضافة عنوان التقرير في المنتصف
        doc.setFontSize(16);
        doc.text(title, doc.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

        // إنشاء الجدول في الـ PDF
        doc.autoTable({
            startY: 25,
            // رؤوس الجدول المعقدة
            head: [
                [{ content: 'عدد المفحوصين', colSpan: 8, styles: { halign: 'center' } }, { content: 'إيجابي بلهارسيا', colSpan: 8, styles: { halign: 'center' } }],
                [{ content: 'عدد المفحوص بول', colSpan: 4, styles: { halign: 'center' } }, { content: 'عدد المفحوص براز', colSpan: 4, styles: { halign: 'center' } }, { content: 'بلهارسيا بولية', colSpan: 4, styles: { halign: 'center' } }, { content: 'بلهارسيا معوية', colSpan: 4, styles: { halign: 'center' } }],
                ['انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر', 'انثى', 'ذكر'],
                ['≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12', '≤12', '>12']
            ],
            // صف البيانات
            body: [dataRow],
            theme: 'grid',
            styles: {
                font: 'Amiri', // تطبيق الخط العربي
                halign: 'center',
                cellPadding: 2
            },
            headStyles: {
                fillColor: [22, 160, 133],
                textColor: 255
            }
        });
        
        // 5. حفظ وتنزيل ملف الـ PDF
        const date = new Date().toISOString().slice(0, 10);
        doc.save(`Lab-Report-${date}.pdf`);

    } catch (error) {
        console.error("حدث خطأ:", error);
        alert("حدث خطأ أثناء إنشاء الملف. يرجى المحاولة مرة أخرى.");
    } finally {
        // إعادة إظهار الزر وإخفاء رسالة التحميل
        generateBtn.style.display = 'block';
        loader.style.display = 'none';
    }
}
